# Référence — Commandes GPIB des appareils (EIP 545 / Racal‑Dana 1996 / Stanford SR620)

> Document de **référence** consolidé (lecture seule, n'altère aucun code).
> Source : `docs/legacy-delphi/Metrologo.ini` + formulaires Delphi `F_ConfigEIP/Racal/Stanford.pas`
> + `U_DeclarationsMETROLOGO.pas`. À utiliser pour câbler le mode « adresses fixes ».

---

## 0. Légende des terminateurs

| Champ | Valeur | Signification |
|------|--------|---------------|
| **WriteTerm** | `0` | NULLEnd — aucune fin d'écriture |
| | `1` | NLEnd — `NL` (LF) ajouté en fin de transfert |
| | `2` | DABEnd — `EOI` positionné en fin de transfert |
| **ReadTerm** | `256` | STOPEnd — fin de lecture sur `EOI` |
| | *autre* | code ASCII du caractère de fin (ex. `10` = `LF`) |
| **TailleHeaderReponse** | `n` | nb de caractères d'entête à ignorer avant le nombre (1 = aucun saut) |

---

## 1. Récapitulatif communication

| Appareil | Section INI | Adresse | WriteTerm | ReadTerm | Header | SRQ |
|----------|-------------|:------:|:---------:|:--------:|:------:|:---:|
| **Stanford SR620** | `[Stanford SR620]` | 16 | 2 (EOI) | 10 (LF) | 1 | non |
| **Racal‑Dana 1996** | `[Racal-Dana 1996]` | 15 | 1 (NL) | 10 (LF) | **3** | **oui** |
| **EIP 545** | `[EIP 545]` | 16 *(14 en commentaire)* | 1 (NL) | 10 (LF) | 1 | **oui** |
| HP 59307A (MUX) | `[HP59307A]` | 12 | 2 (EOI) | 10 (LF) | — | — |

> ⚠️ **Collision d'adresse : Stanford et EIP sont tous deux sur 16.** Indécidable par scan →
> sélection manuelle obligatoire (un seul actif à la fois sur l'adresse 16).

### Voies physiques (dont Voie C / HF)

| Appareil | Voie A | Voie C (HF) | Voie B |
|----------|--------|-------------|--------|
| **Stanford SR620** | 50 Ω / 1 MΩ (`term1,0` / `term1,1`) | UHF `term1,2` | — |
| **Racal‑Dana 1996** | A 50 Ω / 1 MΩ (`FN2 AZ1` / `FN2 AZ0`) | C `FN1` | — |
| **EIP 545** | mono‑voie (bandes `B1/B2/B3`) | — | — |

> ⚠️ Legacy : l'entrée est **un seul sélecteur mutuellement exclusif** (A 50 Ω / A 1 MΩ / C),
> pas des voies indépendantes. Le couplage (AC/DC) ne vaut **que** pour les options de la voie A
> et est masqué sur la voie C / UHF.
>
> **Modélisation WPF (mise à jour)** : le Stanford SR620 n'a **pas** de Voie C physique —
> le passage en UHF est la 3e option du sélecteur d'entrée (`term1,2` voie A, `term2,2`
> voie B, cf. `F_ConfigStanford.pas` `AS_INPUT`), exposée dans les réglages « Entrée
> Voie A/B » (`NbVoies=2`, plus de `ReglagesVoieC`). L'EIP 545 garde son sélecteur de
> bande `B1/B2/B3` (équivalent du `rgpGammesF` Delphi) avec les plages de la doc
> constructeur EIP 545A : Bande 1 = 10 Hz – 100 MHz (1 MΩ, BNC), Bande 2 = 10 MHz – 1 GHz
> (50 Ω, BNC), Bande 3 = 1 – 18 GHz (50 Ω, type N). `ConfEntree` reste `B1` à l'init.

---

## 2. Stanford SR620 — adresse 16

| Étape | Commande |
|-------|----------|
| Initialisation | `*rst;*cls;mode3;autm0;levl1,0` |
| Config entrée (défaut) | `term1,0;tcpl1,1` |
| Déclenche + lit la mesure | `meas?0` *(ancien, commenté : `strt;*wai;xavg?`)* |
| Monocoup | `autm0` |
| SRQ On / Off | *(géré SRQ = non)* |

**Temps de porte (gate) — par slot :**

| Slot | Libellé | Commande |
|:----:|---------|----------|
| 0 | 10 ms | `armm3;size1E0` |
| 1 | 20 ms | `armm3;size2E0` |
| 2 | 50 ms | `armm3;size5E0` |
| 3 | 100 ms | `armm4;size1E0` |
| 4 | 200 ms | `armm4;size2E0` |
| 5 | 500 ms | `armm4;size5E0` |
| 6 | 1 s | `armm5;size1E0` |
| 7 | 2 s | `armm5;size2E0` |
| 8 | 5 s | `armm5;size5E0` |
| 9 | 10 s | `armm5;size1E1` |
| 10 | 20 s | `armm5;size2E1` |
| 11 | 50 s | `armm5;size5E1` |
| 12 | 100 s | `armm5;size1E2` |

**Entrée / impédance / couplage (UI)** — `F_ConfigStanford.pas` + enum `enInputStanford` :

| Réglage | Option | Commande |
|---------|--------|----------|
| Entrée / Impédance | 50 Ω (Low‑Z) | `term1,0` |
| | 1 MΩ | `term1,1` |
| | UHF (voie HF) | `term1,2` |
| Couplage | AC | `tcpl1,1` |
| | DC | `tcpl1,0` |
| *séparateur* | | `;` |

> Le 50 Ω / 1 MΩ **est** la sélection d'entrée (pas de commande d'impédance distincte).
> Couplage **masqué** quand l'entrée = UHF.

---

## 3. Racal‑Dana 1996 — adresse 15

| Étape | Commande |
|-------|----------|
| Initialisation | `QM0 MM1 AT0` |
| Config entrée (défaut) | `FN2 AZ1 AA1` |
| Déclenche + lit la mesure | `RE` |
| Monocoup | `MM1` |
| SRQ On / Off | `QM16` / `QM0` |

**Temps de porte (gate) — par slot :**

| Slot | Libellé | Commande |
|:----:|---------|----------|
| 0 | 10 ms | `GA1E-2` |
| 1 | 20 ms | `GA2E-2` |
| 2 | 50 ms | `GA5E-2` |
| 3 | 100 ms | `GA1E-1` |
| 4 | 200 ms | `GA2E-1` |
| 5 | 500 ms | `GA5E-1` |
| 6 | 1 s | `GA1E0` |
| 7 | 2 s | `GA2E0` |
| 8 | 5 s | `GA5E0` |
| 9 | 10 s | `GA1E1` |
| 10 | 20 s | `GA2E1` |
| 11 | 50 s | `GA5E1` |
| 12 | 100 s | `GA1E2` |

**Entrée / impédance / couplage (UI)** — `F_ConfigRacal.pas` + enum `enInputRacal` :

| Réglage | Option | Commande |
|---------|--------|----------|
| Entrée / Impédance | Entrée A — 50 Ω | `FN2 AZ1` |
| | Entrée A — 1 MΩ | `FN2 AZ0` |
| | Entrée C | `FN1` |
| Couplage | AC | `AA1` |
| | DC | `AA0` |
| *séparateur* | | `␣` (espace) |

> Le 50 Ω / 1 MΩ **est** la sélection d'entrée A (`FN2 AZ1` / `FN2 AZ0`), pas une commande
> d'impédance distincte. Couplage **masqué** quand l'entrée = C.

**Cas spécial — mesure d'intervalle** (`U_DeclarationsMETROLOGO.MesureIntervalleRacalDana`) :

```text
EcritureIEEE 'QM0'                      ; coupe la génération de SRQ
boucle: ReadStatusByte ; si bit16 -> LectureIEEE ; sinon stop  ; vide le buffer
EcritureIEEE 'RE'                       ; relance
SetRemoteLocal(false)
boucle: ReadStatusByte jusqu'à bit16    ; attend mesure prête (Sleep 1s)
LectureIEEE                             ; lit le résultat
```

---

## 4. EIP 545 — adresse 16 (14 en commentaire)

| Étape | Commande |
|-------|----------|
| Initialisation | `HAOPR0FRSR00` |
| Config entrée (défaut) | `B1` |
| Déclenche + lit la mesure | `RS` |
| Monocoup | *(vide)* |
| SRQ On / Off | `SR01` / `SR00` |

**Temps de porte (gate)** — l'EIP ne gère que 3 portes :

| Slot | Libellé | Commande |
|:----:|---------|----------|
| 0 | 10 ms | `R2` |
| 3 | 100 ms | `R1` |
| 6 | 1 s | `R0` |

**Entrée / bande (UI)** — `F_ConfigEIP.pas` + enum `enInputEIP` :

| Réglage | Option | Commande |
|---------|--------|----------|
| Bande de fréquence | Bande 1 | `B1` |
| | Bande 2 | `B2` |
| | Bande 3 | `B3` |

> ⚠️ L'EIP 545 n'a **ni couplage, ni impédance, ni filtre, ni trigger** (le record `TMesure`
> ne porte `Coupling` que pour Stanford et Racal). Seule la bande est réglable.

---

## 5. Multiplexeur HP 59307A — adresse 12

Commandes de voie (`U_DeclarationsMETROLOGO.AS_CDESMUX`, `\n` = LF) :

| Voie | Commande |
|:----:|----------|
| 0 | `A1` |
| 1 | `A2` |
| 2 | `A3` |
| 3 | `A4`␤`B1` |
| 4 | `A4`␤`B2` |
| 5 | `A4`␤`B3` |
| 6 | `A4`␤`B4` |

---

## 6. Format normalisé proposé (1 profil par appareil)

Structure pensée pour coller à `Models/AppareilIEEE.cs` **mais avec une table
gate→commande** (les commandes de gate legacy ne sont pas templatables en `{0}`).

```jsonc
{
  "id": "racal-dana-1996",
  "nom": "Racal-Dana 1996",
  "legacy": true,
  "idnMotif": null,            // null = ne répond pas à *IDN?
  "comm": { "writeTerm": 1, "readTerm": 10, "tailleHeader": 3 },
  "commandes": {
    "init":      "QM0 MM1 AT0",
    "confEntree":"FN2 AZ1 AA1",
    "exeMesure": "RE",
    "monocoup":  "MM1"
  },
  "srq": { "gere": true, "on": "QM16", "off": "QM0" },
  "gates": {                   // clé = slot UI (0=10ms … 12=100s)
    "0":"GA1E-2","1":"GA2E-2","2":"GA5E-2","3":"GA1E-1","4":"GA2E-1",
    "5":"GA5E-1","6":"GA1E0","7":"GA2E0","8":"GA5E0","9":"GA1E1",
    "10":"GA2E1","11":"GA5E1","12":"GA1E2"
  },
  "entrees":  [ {"libelle":"A AZ on","cmd":"FN2 AZ1"},
                {"libelle":"A AZ off","cmd":"FN2 AZ0"},
                {"libelle":"C","cmd":"FN1"} ],
  "couplages":[ {"libelle":"Imp 1","cmd":"AA1"}, {"libelle":"Imp 2","cmd":"AA0"} ],
  "casSpeciaux": { "intervalle": "QM0->pollStatus(bit16)->RE->pollStatus(bit16)->read" }
}
```

> Les profils **Stanford** et **EIP** suivent le même schéma (voir §2 et §4).
> EIP : `gates` ne contient que les slots `0`, `3`, `6`.

---

## 7. Séquence d'envoi commune (pipeline de mesure)

```text
1. SendIFC                              (Interface Clear sur le bus)
2. Config MUX        -> EcritureIEEE(AS_CDESMUX[voie], 12, EOI)
3. Init appareil     -> EcritureIEEE(init,        adresse, writeTerm)   [si pas InitManu]
4. Config entrée     -> EcritureIEEE(confEntree,  adresse, writeTerm)
5. SRQ on            -> EcritureIEEE(srq.on,       adresse, writeTerm)  [si srq.gere]
6. Programme gate    -> EcritureIEEE(gates[slot],  adresse, writeTerm)
7. Mesure            -> EcritureLectureIEEE(exeMesure, ...) puis lecture
8. Parse réponse     -> sauter tailleHeader caractères, convertir en nombre
9. SRQ off           -> EcritureIEEE(srq.off,      adresse, writeTerm)  [si srq.gere]
```

Variantes :
- **Stanford** : pas de SRQ (étapes 5 et 9 sautées).
- **Racal en Interval** : remplace l'étape 7 par la séquence du §3 (cas spécial).
- **EIP** : si la gate demandée n'existe pas (slots ≠ 0/3/6) → refuser ou retomber sur la plus proche.

---

## 8. Idées de câblage (mode « adresses fixes »)

1. **Bonne nouvelle runtime** : `AppareilIEEE.Gates` est déjà un
   `Dictionary<int, GateConfig>` avec **une commande par slot** → le legacy s'y
   mappe **directement** (c'est seulement le *formulaire catalogue* à template `{0}`
   qui ne convient pas). Donc côté exécution, rien à inventer.

2. **Source des profils** : réutiliser `Metrologo.ini` (déjà lu par le chemin legacy)
   ou charger les profils JSON du §6. Indépendant de `*IDN?` → fonctionne pour des
   appareils muets.

3. **Binding de baie** : une liste `{ board, adresse, profilId, actif }`. En mode
   fixe, l'utilisateur ajoute les appareils un par un (EIP | Racal | Stanford + adresse).
   Le `profilId` choisi détermine le jeu de commandes appelé.

4. **Exclusion mutuelle d'adresse** : interdire deux appareils *actifs* sur la même
   adresse (EIP/Stanford = 16). Au lancement d'une mesure, un seul est sélectionné.

5. **Persistance** : mémoriser le binding par poste (Baie) pour ne pas re‑saisir à
   chaque session ; la Paillasse peut rester en scan.

6. **Compat type de mesure** : conserver les cas métier existants (Racal Interval
   `QM0…RE`, EIP gates limitées, Stanford sans SRQ).
