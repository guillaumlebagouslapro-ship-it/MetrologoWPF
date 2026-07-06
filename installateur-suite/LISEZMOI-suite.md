# Installateur « Suite ASERTI » (cases à cocher)

Un seul `Suite-ASERTI-Setup.exe` qui propose un **assistant avec cases à cocher** :
on choisit d'installer **Asertools**, **Metrologo**, ou les deux.

## Comment ça marche

L'installateur suite **n'embarque pas** les applications. Pour chaque case cochée,
il **lance le `Setup.exe` Velopack** correspondant, déjà déposé sur le réseau `M:`.

Conséquences :
- L'installateur suite est **minuscule** (~2 Mo).
- Il installe **toujours la dernière version** publiée (il lance le Setup courant sur `M:`).
- Chaque application garde sa **mise à jour automatique indépendante** ensuite.

```
Suite-ASERTI-Setup.exe
   ├─ [x] Asertools  → lance M:\exe_spe\Data_Metrologo\Asertools\Asertools-win-Setup.exe
   └─ [x] Metrologo  → lance M:\exe_spe\Data_Metrologo\Metrologo\Metrologo-win-Setup.exe
```

## Prérequis

- Les deux applications doivent **déjà être publiées** sur `M:` (via leurs scripts
  `publier-asertools.bat` / `publier-metrologo.bat`). Si un `Setup.exe` manque, le
  composant est simplement ignoré avec un avertissement.
- Pour **construire** la suite : **Inno Setup 6** installé sur le poste de build
  (`winget install JRSoftware.InnoSetup`).

## Construire / mettre à jour la suite (dev, au bureau, `M:` connecté)

1. Double-clic sur **`construire-suite.bat`** (ce dossier).
2. Il produit **`M:\exe_spe\Data_Metrologo\Suite\Suite-ASERTI-Setup.exe`**.

À refaire uniquement si tu modifies `suite-aserti.iss` (ex. ajouter un 3ᵉ logiciel).
Publier une nouvelle version d'une app **ne nécessite pas** de reconstruire la suite.

## Installer sur un poste (utilisateur)

1. Connecter le lecteur `M:` (raccourci du bureau).
2. Lancer **`M:\exe_spe\Data_Metrologo\Suite\Suite-ASERTI-Setup.exe`**.
3. Choisir **Installation personnalisée** et cocher ce qu'on veut (ou **complète**
   pour les deux), puis **Suivant**.
4. Chaque application cochée s'installe. Ensuite, elles se mettent à jour toutes
   seules au lancement.

## Fichiers

| Fichier | Rôle |
|---|---|
| `suite-aserti.iss` | Script Inno Setup (composants, logique) |
| `construire-suite.bat` | Compile le script → `Suite-ASERTI-Setup.exe` sur `M:` |

## Ajouter un 3ᵉ logiciel plus tard

Dans `suite-aserti.iss` : ajouter une ligne `[Components]` et une ligne `[Run]`
pointant vers le `Setup.exe` du nouveau logiciel sur `M:`, puis reconstruire.
