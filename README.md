# 📅 Planning → Agenda

Convertit l'export Excel (`.xlsx`) d'un planning annuel issu du logiciel RH en fichier **iCalendar (`.ics`)**, importable dans n'importe quel agenda : Google Agenda, Apple Agenda (iCloud), Outlook…

- **Une seule page HTML**, hébergée gratuitement sur GitHub Pages, utilisable sur ordinateur et mobile.
- **Aucune donnée envoyée sur internet** : le fichier Excel est lu et converti entièrement dans le navigateur.
- **Contrôle avant export** : heures recalculées et comparées aux totaux du logiciel RH, congés listés, confirmation obligatoire.

---

## Sommaire

1. [Fonctionnement](#1-fonctionnement)
2. [Format d'export attendu](#2-format-dexport-attendu)
3. [Architecture du code](#3-architecture-du-code)
4. [Logique métier détaillée](#4-logique-métier-détaillée)
5. [Génération du fichier ICS](#5-génération-du-fichier-ics)
6. [Configuration](#6-configuration)
7. [Déploiement (GitHub Pages)](#7-déploiement-github-pages)
8. [Développement et tests](#8-développement-et-tests)
9. [Maintenance : tâches courantes](#9-maintenance--tâches-courantes)
10. [Dépannage](#10-dépannage)
11. [Confidentialité et sécurité](#11-confidentialité-et-sécurité)
12. [Limites connues](#12-limites-connues)
13. [Historique des versions](#13-historique-des-versions)

---

## 1. Fonctionnement

Le parcours utilisateur suit 5 étapes visibles sur la page :

| Étape | Rôle |
|---|---|
| **1. Export du planning** | Sélection (ou glisser-déposer) du `.xlsx` exporté du logiciel RH. |
| **2. Codes à exporter** | Liste des codes rencontrés (J1, N1, CA…) : case à cocher pour les inclure, titre personnalisable dans l'agenda. |
| **3. Options** | Période « Du / Au », préfixe des titres, rappel, regroupement des absences consécutives. |
| **4. Vérification** | Chiffres clés, heures par mois comparées au logiciel RH, congés posés, confirmation obligatoire. |
| **5. Aperçu et export** | Liste des événements, bouton **Télécharger** (et **Partager…** sur mobile). |

En bas de page :
- un **guide d'import** par plateforme (iPhone, Mac, Google, Android, Outlook, partage, dépannage) ;
- un lien discret **« Problème technique ? »** qui ouvre un parcours de confirmation en deux questions, puis un formulaire de signalement GitHub (issue) prérempli.

```
.xlsx ──► SheetJS ──► parseWorkbook() ──► summarize()  ──► Vérification (étape 4)
                              │
                              └────────► buildEvents() ──► toICS() ──► planning_AAAA-MM-JJ_AAAA-MM-JJ.ics
```

---

## 2. Format d'export attendu

Le classeur exporté par le logiciel RH contient plusieurs onglets. **Seuls les onglets suivants sont utilisés** (repérés par leur nom, sans tenir compte de la casse ni des accents) :

| Onglet | Mots recherchés dans le nom | Obligatoire | Usage |
|---|---|---|---|
| `Planning codes` | `planning` + `code` | ✅ | Code du jour (J1, N1, CA, RH…) |
| `Planning plages` | `planning` + `plage` | recommandé | Horaires réels du jour (`07:00/19:00`) |
| `Planning durées` | `planning` + `dur` | recommandé | Durée travaillée du jour (`12:00`), utilisée pour le contrôle des heures |
| `Détail codes` | `detail` + `code` | recommandé | Libellé, plage et durée par défaut de chaque code |

Les onglets `Astreintes`, `Employé reçu` et `Employé prêté` sont ignorés.

### 2.1 Structure des onglets « Planning … »

Les trois onglets « Planning » ont la même grille annuelle :

```
Ligne 1 :          Planning annuel du 01/01/26 au 31/12/26 de NOM PRÉNOM    Edité le …
Ligne 2 : janvier  │        │ février │        │ mars  │ …     ← 1 bloc de 3 colonnes par mois
Ligne 3 : J 01     │ RH     │ D 01    │ RH     │ D 01  │ N1
Ligne 4 : V 02     │ RH     │ L 02    │ CA     │ L 02  │ N1
…
Ligne 34:          │ 180h00 │         │ 120h00 │       │ 192h00  … 1708h00   ← totaux mensuels + total général
```

- **Colonne 1 du bloc** : lettre du jour + numéro (`J 01`, `V 02`…).
- **Colonne 2 du bloc** : valeur du jour (code, plage horaire ou durée selon l'onglet).
- **Titre** (ligne 1) : fournit l'**année de début** (`du 01/01/26`) et le **nom de l'agent**.
- **Ligne de totaux** sous la grille : total d'heures par mois (`180h00`) et total général (`1708h00`).

> Le code ne dépend **pas** de numéros de lignes ou de colonnes fixes : il cherche la ligne contenant au moins deux noms de mois, puis lit chaque bloc. Un planning à cheval sur deux années (septembre → août) est géré : l'année est incrémentée quand le mois « recule ».

### 2.2 Onglet « Détail codes »

```
Codes horaire                          Codes événement           Codes lieux
Code │ Nom                  │ Plage        │ Durée │   │ Code │ Nom          │   │ Code │ Nom
RH   │ Repos Hebdo          │              │ 00:00 │   │ CA   │ ANNUEL       │   │ NER  │ NEO / REANIMATION
J1   │ J12 de 07h00 à 19h00 │ 07:00/19:00  │ 12:00 │   │ FO   │ FORMATION    │   │ …
N1   │ Nuit de 19h00 à 07h00│ 19:00/07:00  │ 12:00 │   │ EX   │ EXCEPTIONNEL │
-    │ Journée non travaillée│             │ 00:00 │   │ …
Jq3  │ J4 de 08h00 à 12h00  │ 08:00/12:00  │ 04:00 │
```

Chaque tableau est repéré par un en-tête `Code │ Nom`. Sa **catégorie** (`horaire`, `evenement`, `lieu`) est déduite du libellé situé juste au-dessus (`Codes horaire`, `Codes événement`, `Codes lieux`).

---

## 3. Architecture du code

Tout tient dans **`index.html`** (aucune étape de build, aucune dépendance à installer).

```
index.html
├── <head>
│   ├── <script src="…/xlsx.full.min.js">     ← SheetJS (lecture Excel), chargé depuis le CDN officiel
│   └── <style>                               ← tout le CSS (variables de couleur dans :root)
├── <main>
│   ├── #drop / #file                         ← étape 1
│   ├── #secCodes                             ← étape 2
│   ├── #secOpts                              ← étape 3
│   ├── #secCheck                             ← étape 4 (vérification)
│   ├── #secPreview                           ← étape 5 (aperçu + export)
│   └── #secHelp                              ← guide d'import
├── <footer> #btnIssue + <dialog id="issueDlg">  ← « Problème technique ? »
├── <script id="core">                        ← LOGIQUE PURE, sans DOM (testable sous Node)
└── <script>                                  ← INTERFACE (DOM, événements, localStorage)
```

### 3.1 Séparation « core » / « interface »

Le bloc `<script id="core">` ne touche **jamais** au DOM. Il peut donc être extrait et exécuté sous Node pour les tests (voir [§8](#8-développement-et-tests)). Il se termine par :

```js
if (typeof module !== 'undefined') module.exports = { parseWorkbook, buildEvents, toICS, summarize };
```

| Fonction (core) | Rôle |
|---|---|
| `findSheet(wb, words)` | Trouve un onglet dont le nom contient tous les mots (sans accents). |
| `rows(XLSX, ws)` | Convertit un onglet en tableau 2D (valeurs texte formatées). |
| `readGrid(grid, startYear)` | Lit une grille annuelle → `{ cells:[{date, value}], warnings, monthTotals, grandTotal }`. |
| `readCodes(grid)` | Lit « Détail codes » → `{ J1: {nom, plage, duree, cat}, … }`. |
| `parseHM(v)` | `"12:00"` ou `"180h00"` → minutes (ou `null`). |
| `parsePlages(s)` | `"07:00/19:00"` → `[{sh:7, sm:0, eh:19, em:0}]` (plusieurs plages possibles). |
| `parseWorkbook(XLSX, wb)` | Point d'entrée : assemble tout → objet `parsed`. |
| `defaultTitle(u)` | Titre par défaut d'un code dans l'agenda. |
| `summarize(parsed, cfg)` | Chiffres de l'étape « Vérification ». |
| `buildEvents(parsed, cfg)` | Liste des événements à exporter selon les réglages. |
| `toICS(events, cfg)` | Sérialise en texte iCalendar (RFC 5545). |
| `fold(line)`, `esc(s)` | Pliage des lignes à 75 octets, échappement des caractères spéciaux. |

| Fonction (interface) | Rôle |
|---|---|
| `handleFile(file)` | Lecture du fichier, gestion des erreurs, affichage des sections. |
| `renderCodes()` | Tableau des codes (étape 2), avec les choix mémorisés. |
| `readCfg()` | Lit tous les réglages du formulaire → objet `cfg`. |
| `refresh()` | Recalcule tout à chaque modification (événements, ICS, vérification, aperçu). |
| `renderCheck(cfg)` | Affiche l'étape « Vérification ». |
| `updateLock()` | Active le téléchargement seulement si la confirmation est cochée. |
| `renderPreview(cfg)` | Liste des événements (étape 5). |
| `download()` / `share()` | Téléchargement du `.ics` / partage natif (mobile) avec repli sur le téléchargement. |
| `githubRepo()` / `issueUrl()` | Détection du dépôt et construction du lien de signalement. |

### 3.2 Structures de données

```js
// Résultat de parseWorkbook()
parsed = {
  person: 'NOM PRÉNOM',            // extrait du titre (n'est jamais exporté dans un signalement)
  first: '2026-01-01', last: '2026-11-30',
  days: [ { date: '2026-01-03', code: 'J1', plages: [{sh:7,sm:0,eh:19,em:0}], minutes: 720 }, … ],
  used: { J1: { code, count, info, timed, kind: 'work'|'abs'|'rest', defaultInclude, defaultTitle }, … },
  monthTotals: { '2026-01': 10800, … },  // minutes, lues sous la grille
  grandTotal: 102480,                    // minutes
  warnings: [],                          // incohérences jour de semaine / date
};

// Réglages lus dans le formulaire (readCfg)
cfg = {
  codes: { J1: { include: true, title: 'Jour 12h' }, RH: { include: false, title: 'Repos hebdo' }, … },
  from: '2026-11-07', to: '2026-11-30',
  prefix: 'Lise · ', alarm: '60', merge: true, uidKey: 'NOM PRÉNOM',
};
```

---

## 4. Logique métier détaillée

### 4.1 Classement des codes

Chaque code rencontré reçoit un **type** (`kind`), qui détermine s'il est coché par défaut :

```js
u.kind = u.timed ? 'work'                                              // a au moins une plage horaire
       : (u.info && u.info.cat === 'horaire') || u.code === '-' ? 'rest' // code horaire sans plage (RH, -)
       : 'abs';                                                         // code événement (CA, FO, EX…)
u.defaultInclude = u.kind !== 'rest';
```

| Type | Exemples | Export | Coché par défaut |
|---|---|---|---|
| `work` | J1, N1, Jq3 | Événement **horaire** | ✅ |
| `abs` | CA, FO, EX | Événement **journée entière** | ✅ |
| `rest` | RH, `-` | Événement journée entière | ❌ |

### 4.2 Horaires et durées

- **Horaires** : on lit d'abord l'onglet `Planning plages` pour le jour (cela gère un horaire exceptionnel), sinon la plage par défaut du code dans `Détail codes`.
- **Fin ≤ début** ⇒ le poste finit le **lendemain** (nuits N1 : 19:00 → 07:00 +1).
- **Durée** (contrôle des heures) : onglet `Planning durées` en priorité. Sinon, longueur des plages.

```js
let minutes = durByDate[c.date] ?? null;
if (minutes === null)
  minutes = plages.reduce((s, p) => s + (((p.eh*60+p.em) - (p.sh*60+p.sm) + 1440) % 1440 || 1440), 0);
```

### 4.3 Titres par défaut

Définis dans `defaultTitle()` : `Nuit` pour un libellé commençant par « Nuit », `Jour 12h` d'après la durée, sinon le libellé du code mis en forme. Les libellés en majuscules du logiciel RH sont traduits via le dictionnaire `NICE` :

```js
const NICE = { 'ANNUEL': 'Congé annuel', 'FORMATION': 'Formation', 'RECUPERATION': 'Récupération',
               'EXCEPTIONNEL': 'Absence exceptionnelle', 'HS A RECUPERER': 'Heures sup. à récupérer',
               'FORMATION A RECUPERER': 'Formation à récupérer' };
```

L'utilisateur peut modifier chaque titre. Ses choix sont mémorisés dans le navigateur ([§6.4](#64-données-mémorisées-dans-le-navigateur)).

### 4.4 Vérification (étape 4)

`summarize(parsed, cfg)` calcule :

| Contrôle | Calcul | Affichage |
|---|---|---|
| Période | `cfg.from` → `cfg.to`, nombre de jours | tuile |
| Heures sur la période | somme des `minutes` des jours `work` dans la période | tuile |
| Heures par mois | somme des `minutes` de **tous** les jours du mois (fichier entier) | tableau, comparé à `monthTotals` (✓ ou ≠) |
| Total fichier | somme de tous les jours | comparé à `grandTotal` |
| Congés et absences | jours `abs` dans la période, regroupés en périodes consécutives | liste |
| Codes décochés | heures exportées ≠ heures de la période | message ℹ️ |

Le bouton de téléchargement reste **désactivé** tant que la case « J'ai vérifié… » n'est pas cochée. Toute modification de la période **décoche** la confirmation.

> ✅ Validé sur l'export de référence 2026 : 11 mois sur 11 et le total de 1 708 h correspondent aux totaux du logiciel RH.

---

## 5. Génération du fichier ICS

### 5.1 Exemple de sortie

```ics
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//planning-ics//FR
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VTIMEZONE
TZID:Pacific/Tahiti
…
END:VTIMEZONE
BEGIN:VEVENT
UID:20260210-N1-0-nomprenom@planning-ics
DTSTAMP:20261107T214900Z
DTSTART;TZID=Pacific/Tahiti:20260210T190000
DTEND;TZID=Pacific/Tahiti:20260211T070000
SUMMARY:Nuit
DESCRIPTION:Code planning : N1
CATEGORIES:N1
END:VEVENT
BEGIN:VEVENT
UID:20260202-CA-nomprenom@planning-ics
DTSTAMP:20261107T214900Z
DTSTART;VALUE=DATE:20260202
DTEND;VALUE=DATE:20260208
TRANSP:TRANSPARENT
SUMMARY:Congé annuel
DESCRIPTION:Code planning : CA
CATEGORIES:CA
END:VEVENT
END:VCALENDAR
```

### 5.2 Règles de conformité (RFC 5545)

| Règle | Implémentation |
|---|---|
| Fins de ligne `CRLF` | `L.map(fold).join('\r\n') + '\r\n'` |
| Lignes ≤ 75 octets | `fold()` coupe en octets UTF-8 (accents compris) et continue par un espace |
| Échappement | `esc()` : `\` `;` `,` et retours à la ligne |
| Événement journée | `DTSTART;VALUE=DATE` + `DTEND` **exclusif** (lendemain du dernier jour) |
| Absences | `TRANSP:TRANSPARENT` (n'apparaissent pas « occupé ») |
| Fuseau | imposé `Pacific/Tahiti` (UTC−10) : bloc `VTIMEZONE` inclus + `TZID` sur chaque horaire |
| Rappel | `VALARM` (`TRIGGER:-PT{n}M`) uniquement sur les postes horaires |

### 5.3 UID stables

```
UID = AAAAMMJJ-<code>-<n°plage>-<nom agent>@planning-ics   (postes)
UID = AAAAMMJJ-<code>-<nom agent>@planning-ics             (journées)
```

Le même jour et le même code produisent toujours le même UID : certains agendas (Apple notamment) peuvent alors mettre à jour l'événement au lieu de le dupliquer. **Ne modifiez pas ce format** sans nécessité, car les imports précédents ne seraient plus reconnus.

---

## 6. Configuration

### 6.1 Dépôt GitHub des signalements

```js
// index.html, bloc « Signalement d'un problème technique »
const GITHUB_REPO = '';   // ex. 'mon-compte/planning-ics'
```

Laissé vide, le dépôt est **détecté automatiquement** depuis l'adresse GitHub Pages (`https://<propriétaire>.github.io/<dépôt>/` → `<propriétaire>/<dépôt>`). Renseignez-le seulement si la page est hébergée ailleurs, ou si les signalements doivent aller dans un autre dépôt.

Le lien ouvre `https://github.com/<dépôt>/issues/new?title=…&body=…&labels=bug`. Le corps contient un modèle et des informations techniques (navigateur, écran, fichier lu ou non, dernière erreur), **jamais le planning ni le nom de l'agent**.

> 🔔 Pour recevoir les notifications : dépôt → **Watch** → *All activity* (ou au minimum *Issues*), et vérifiez vos réglages sur <https://github.com/settings/notifications>. Le label `bug` n'est appliqué que si l'auteur a les droits sur le dépôt ; sinon, GitHub l'ignore sans erreur.

### 6.2 Fuseau horaire

Le fuseau est **imposé** : `Pacific/Tahiti` (Polynésie française, UTC−10, sans heure d'été). Il est défini par les constantes `TZ` et `VTZ` du bloc core :

```js
const TZ = 'Pacific/Tahiti';
const VTZ = ['BEGIN:VTIMEZONE', `TZID:${TZ}`, /* STANDARD -1000 */ 'END:VTIMEZONE'];
```

Pour changer de fuseau, modifiez ces deux constantes (un fuseau avec heure d'été nécessite un bloc `DAYLIGHT`/`STANDARD` complet).

### 6.3 Bibliothèque SheetJS

```html
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
```

- Utilisez le **CDN officiel** `cdn.sheetjs.com`. Les versions npm au-delà de 0.18.5 ne sont plus publiées sur le registre npm.
- Pour une **utilisation hors ligne** ou pour ne dépendre d'aucun CDN : téléchargez `xlsx.full.min.js` dans le dépôt et remplacez le `src` par `./xlsx.full.min.js`.
- Après une **mise à jour de version**, relancez les tests ([§8](#8-développement-et-tests)).

### 6.4 Données mémorisées dans le navigateur

| Clé `localStorage` | Contenu |
|---|---|
| `planning-ics:codes` | `{ J1: { include, title }, … }` : choix de l'étape 2 |
| `planning-ics:opts` | `{ prefix, alarm, merge }` : options de l'étape 3 |

Les données ne sont mémorisées que sur l'appareil, et l'accès est protégé par `try/catch` (la navigation privée peut le bloquer). **Le planning lui-même n'est jamais stocké.** Pour réinitialiser :

```js
// Console du navigateur (F12), sur la page de l'application
Object.keys(localStorage).filter(k => k.startsWith('planning-ics:')).forEach(k => localStorage.removeItem(k));
```

### 6.5 Apparence

La charte graphique reprend la palette du **CHPF** (Centre Hospitalier de la Polynésie française). Les couleurs sont des variables CSS dans `:root` : d'abord les couleurs de la palette (`--chpf-*`), puis les rôles utilisés par la page :

```css
:root {
  --chpf-ardoise: #3e444a; --chpf-vert: #00ae5a; --chpf-bleu: #22bffd; --chpf-orange: #ffa837; --chpf-corail: #ff5151;
  /* … */
  --accent: var(--chpf-ardoise);   /* bandeau, titres, boutons */
  --accent2: var(--chpf-vert);     /* éléments décoratifs */
  --bg: var(--chpf-gris-50);       /* fond */
  --ok: #00753d; --warn: #8f5600; --err: #c62828;  /* variantes foncées pour le texte */
}
```

Les couleurs vives du logo (vert, bleu, orange, corail) n'atteignent pas un contraste suffisant pour du texte (AA, 4,5:1). Elles sont donc réservées aux aplats et aux décors (liseré quadricolore sous le bandeau, bordures). Pour le texte coloré, la page utilise des variantes foncées de ces teintes.

---

## 7. Déploiement (GitHub Pages)

### Première mise en ligne

1. Créez un dépôt **public**, par exemple `planning-ics`.
2. Déposez `index.html` (et ce `README.md`) à la racine.
3. **Settings → Pages → Build and deployment** : *Source* = *Deploy from a branch*, *Branch* = `main`, dossier `/ (root)` → **Save**.
4. Au bout d'une à deux minutes, l'URL `https://<votre-nom>.github.io/planning-ics/` s'affiche dans **Settings → Pages**. L'onglet **Actions** montre l'état du déploiement.

### Mise à jour

Via l'interface : ouvrez `index.html` → ✏️ → collez la nouvelle version → **Commit changes**.

Ou en ligne de commande :

```bash
git clone https://github.com/<votre-nom>/planning-ics.git
cd planning-ics
# … modifier index.html …
git add index.html
git commit -m "Correction : lecture des totaux mensuels"
git push
```

Le site se met à jour automatiquement en environ une minute. Côté utilisateur, forcez le rechargement (`Ctrl+F5` / `Cmd+Shift+R`) en cas de cache.

### Arborescence recommandée

```
planning-ics/
├── index.html          ← l'application (seul fichier indispensable)
├── README.md           ← ce document
├── tests/
│   └── test-core.js    ← test automatisé (voir §8)
└── .gitignore          ← exclure tout export réel
```

```gitignore
# .gitignore — ne jamais versionner de planning réel
*.xlsx
*.xls
*.ics
```

---

## 8. Développement et tests

### 8.1 Lancer la page en local

Ouvrir `index.html` dans le navigateur suffit (double-clic). Pour reproduire exactement les conditions d'un serveur web :

```bash
python3 -m http.server 8000
# puis ouvrir http://localhost:8000
```

> En local, le lien « Contacter l'administrateur » est remplacé par un message : c'est normal, le dépôt n'est détecté que sur `*.github.io` (ou via `GITHUB_REPO`).

### 8.2 Test automatisé de la logique (Node.js)

Le script extrait le bloc `<script id="core">` de `index.html`, l'exécute sous Node et vérifie les résultats sur un export **anonymisé** (à conserver hors du dépôt ou anonymisé).

```bash
npm install xlsx@0.18.5          # version npm suffisante pour les tests
node tests/test-core.js chemin/vers/export_anonymise.xlsx
```

```js
// tests/test-core.js
const fs = require('fs');
const assert = require('assert');
const XLSX = require('xlsx');

// 1. Charger la logique « core » depuis index.html
const html = fs.readFileSync(__dirname + '/../index.html', 'utf8');
const core = html.match(/<script id="core">([\s\S]*?)<\/script>/)[1];
const mod = { exports: {} };
new Function('module', 'TextEncoder', core)(mod, TextEncoder);
const { parseWorkbook, buildEvents, toICS, summarize } = mod.exports;

// 2. Lire l'export
const wb = XLSX.read(fs.readFileSync(process.argv[2]), { type: 'buffer' });
const p = parseWorkbook(XLSX, wb);
const codes = Object.fromEntries(Object.values(p.used).map(u => [u.code, { include: u.defaultInclude, title: u.defaultTitle }]));
const cfg = { codes, from: p.first, to: p.last, merge: true, uidKey: 'test' };

// 3. Contrôles
assert.strictEqual(p.warnings.length, 0, 'Jours de semaine incohérents : ' + p.warnings.join(', '));
const s = summarize(p, cfg);
s.months.forEach(m => assert.notStrictEqual(m.match, false, `Écart d'heures en ${m.key} : ${m.minutes/60}h vs ${m.rh/60}h`));
assert.notStrictEqual(s.grandMatch, false, 'Total général différent du logiciel RH');

const ev = buildEvents(p, cfg);
const night = ev.find(e => e.type === 'timed' && e.end.date !== e.start.date);
if (night) assert.ok(night.end.date > night.start.date, 'Une nuit doit finir le lendemain');

const ics = toICS(ev, cfg);
assert.ok(ics.startsWith('BEGIN:VCALENDAR\r\n') && ics.endsWith('END:VCALENDAR\r\n'), 'Enveloppe ICS');
assert.ok(ics.split('\r\n').every(l => Buffer.byteLength(l) <= 75), 'Lignes > 75 octets');

console.log(`✅ ${p.days.length} jours, ${ev.length} événements, ${s.fileMinutes / 60} h — tous les contrôles passent`);
```

### 8.3 Valider un fichier ICS produit

```bash
pip install icalendar
python3 -c "
from icalendar import Calendar
c = Calendar.from_ical(open('planning.ics','rb').read())
ev = list(c.walk('VEVENT')); print(len(ev), 'événements')
for e in ev[:5]: print(e['SUMMARY'], e['DTSTART'].dt, '→', e['DTEND'].dt)
"
```

Vous pouvez aussi utiliser un validateur en ligne (icalendar.org), avec un fichier **anonymisé**.

### 8.4 Checklist de recette avant mise en ligne

- [ ] Le test `tests/test-core.js` passe sur un export récent.
- [ ] Ordinateur : charger l'export → les 5 étapes s'affichent, aucune erreur dans la console (F12).
- [ ] La vérification affiche ✓ sur tous les mois.
- [ ] Le téléchargement est bloqué tant que la confirmation n'est pas cochée.
- [ ] Le `.ics` s'importe dans un agenda **de test** (Google sur ordinateur et iPhone).
- [ ] Les nuits apparaissent de 19:00 à 07:00 le lendemain, les congés en journée entière.
- [ ] Mobile : la page est lisible et « Partager… » ouvre le menu de partage ou télécharge le fichier.
- [ ] « Problème technique ? » → Oui → Oui → le lien ouvre bien le formulaire d'issue du bon dépôt.

---

## 9. Maintenance : tâches courantes

### Le logiciel RH ajoute un nouveau code

Rien à faire : un code inconnu apparaît automatiquement à l'étape 2. S'il a une plage dans `Détail codes` ou `Planning plages`, il est classé « Travail ». Sinon, il est classé « Absence », ou « Repos » s'il figure parmi les codes horaires. Pour lui donner un joli titre par défaut, complétez `NICE` :

```js
const NICE = { …, 'TELETRAVAIL': 'Télétravail' };   // clé = libellé « Nom » en MAJUSCULES dans Détail codes
```

### Le format d'export change

| Symptôme | Où regarder |
|---|---|
| « Onglet « Planning codes » introuvable » | `findSheet()` et les mots-clés dans `parseWorkbook()` |
| « Ligne des mois introuvable » | `MONTHS` (nouveau libellé, autre langue) |
| 0 jour lu | `DAY_RE` (format du libellé de jour, ex. `Lun 01` au lieu de `L 01`) |
| Horaires absents | `PLAGE_RE` (séparateur différent, ex. `07h00-19h00`) |
| Mauvaise année | regex du titre dans `parseWorkbook()` (`du JJ/MM/AA au …`) |
| Totaux RH « — » | lecture des totaux en fin de `readGrid()` (format `180h00`) |

Exemple : accepter aussi `07h00-19h00` comme séparateur de plage :

```js
const PLAGE_RE = /(\d{1,2})[:hH](\d{2})\s*[\/\-–]\s*(\d{1,2})[:hH](\d{2})/g;
```

### Changer le comportement par défaut d'un type

```js
// parseWorkbook() — ex. exporter aussi les repos par défaut
u.defaultInclude = true;
```

### Ajouter une colonne d'information dans l'événement

Les champs ICS sont écrits dans `toICS()`. Exemple pour ajouter un lieu fixe :

```js
L.push(`SUMMARY:${esc(e.title)}`, `LOCATION:${esc('CHU – Néonatologie')}`, …);
```

### Mettre à jour SheetJS

1. Consultez la dernière version sur <https://cdn.sheetjs.com/>.
2. Remplacez `xlsx-0.20.3` dans la balise `<script src>`.
3. Lancez le test ([§8.2](#82-test-automatisé-de-la-logique-nodejs)) et la recette ([§8.4](#84-checklist-de-recette-avant-mise-en-ligne)).

### Traiter un signalement (issue)

1. Lisez les **informations techniques** repliées en bas de l'issue (navigateur, fichier lu, dernière erreur).
2. Reproduisez le problème avec un export anonymisé.
3. Corrigez le problème, puis ajoutez si possible le cas au test.
4. Mentionnez `Fixes #<n°>` dans le message de commit pour fermer l'issue automatiquement.

---

## 10. Dépannage

| Problème | Cause probable | Solution |
|---|---|---|
| « La bibliothèque de lecture Excel n'a pas pu se charger » | Pas de connexion, CDN bloqué (réseau d'entreprise) | Réessayer sur un autre réseau, ou héberger `xlsx.full.min.js` dans le dépôt ([§6.3](#63-bibliothèque-sheetjs)) |
| Jours décalés d'un jour | Fichier issu d'un autre format de date | Regarder les avertissements « jour de semaine incohérent » affichés à l'étape 1 |
| Heures ≠ logiciel RH | Durée absente de `Planning durées`, ou nouveau code | Vérifier l'onglet durées, comparer jour par jour dans l'aperçu |
| Doublons dans l'agenda | Fichier importé deux fois | Importer par périodes disjointes (« Du / Au »), ou supprimer l'agenda dédié et réimporter |
| Horaires décalés de plusieurs heures | Fuseau de l'agenda différent | Le fichier est à l'heure de Tahiti (UTC−10) : vérifier le fuseau de l'agenda |
| « Partager… » absent | Ordinateur, ou navigateur sans partage de fichiers | Normal : utiliser « Télécharger » |
| Lien de signalement inactif | Page ouverte hors GitHub Pages | Renseigner `GITHUB_REPO` ([§6.1](#61-dépôt-github-des-signalements)) |

---

## 11. Confidentialité et sécurité

- **Traitement 100 % local** : le `.xlsx` est lu par `FileReader`/`arrayBuffer()` dans le navigateur. Aucune requête réseau ne transporte les données (seul SheetJS est téléchargé depuis son CDN).
- **Dépôt public** : le code est visible de tous, **jamais les plannings**. Ne versionnez aucun export réel (voir `.gitignore`).
- **Signalements GitHub** : les issues d'un dépôt public sont publiques. Le formulaire prérempli exclut volontairement le nom de l'agent et le contenu du planning, et l'utilisateur est averti de ne rien y coller de personnel.
- **Affichage** : les valeurs issues du fichier sont insérées via `textContent` (pas d'injection HTML).
- **Pour renforcer** : héberger SheetJS localement et ajouter une politique CSP :

```html
<meta http-equiv="Content-Security-Policy"
      content="default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'none'">
```

---

## 12. Limites connues

- **Import ≠ synchronisation** : un `.ics` importé est une copie. Les changements ultérieurs du planning doivent être réimportés (par période) ou corrigés à la main dans l'agenda.
- **Google Agenda mobile** ne sait pas importer de fichier : l'import se fait sur ordinateur.
- **Congés comptés en jours calendaires** (week-ends inclus) dans la vérification, ce qui peut différer d'un décompte en jours ouvrés.
- **Fuseau** : imposé à `Pacific/Tahiti` (voir [§6.2](#62-fuseau-horaire) pour le modifier).
- **Un agent par fichier** : l'export doit concerner une seule personne.
- **Signalement** : nécessite un compte GitHub.

---

## 13. Historique des versions

| Version | Date | Changements |
|---|---|---|
| 1.2 | 2026-11-07 | Fuseau imposé `Pacific/Tahiti` (suppression de l'option), charte graphique CHPF, fenêtre « Avertissement » à l'ouverture, numéro de version en bas de page (joint aux signalements). |
| 1.0 | 2026-11-07 | Première version : lecture de l'export, choix des codes, options, vérification des heures et congés avec confirmation, export ICS, partage mobile, guide d'import, signalement GitHub. |

> Pensez à compléter ce tableau à chaque mise en ligne, et à mettre à jour la constante `APP_VERSION` dans `index.html` (affichée en bas de page et jointe aux signalements).
