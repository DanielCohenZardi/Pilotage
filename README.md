# Export brut des balances comptables Pennylane

Génère automatiquement, chaque mois, **un unique classeur Excel** avec la
balance comptable de chaque mois disponible, pour une ou plusieurs sociétés
Pennylane, commité dans un dépôt GitHub **privé**.

**Coût : 0 €.** Tout tourne sur l'offre gratuite de GitHub Actions.

> Pas de page web publique ici : comme ce sont des données financières, le
> classeur reste dans le dépôt privé et se télécharge directement depuis
> GitHub (voir plus bas). GitHub Pages sur un dépôt privé nécessite un
> abonnement payant (GitHub Pro) — inutile pour ce cas d'usage.

## Principe

Deux endpoints Pennylane, aucun autre :

- [`fiscal_years`](https://pennylane.readme.io/reference/company-fiscal-years)
  → détermine la plage de mois disponible (du début du plus ancien exercice
  fiscal jusqu'au mois en cours).
- [`trial_balance`](https://pennylane.readme.io/reference/gettrialbalance)
  → récupère la balance comptable, mois par mois, sur toute cette plage.

Résultat : **un seul fichier Excel** (`reports/balances_dernier.xlsx`), une
ligne par compte, par mois, par société :

| Société | Mois | Compte | Libellé | Solde |
|---|---|---|---|---|
| Ma Société A | 2026-08 | 60600000 | Achats non stockés | 670.30 |
| Ma Société A | 2026-08 | 70600000 | Prestations de services | -2300.50 |
| Ma Société B | 2026-08 | 51200000 | Banque | 1200.00 |

**"Solde" = Débit − Crédit du mois** (mouvement net de la période, pas un
cumul depuis le début de l'exercice). Données volontairement "brutes",
prêtes pour un tableau croisé dynamique — aucun indicateur calculé pour
l'instant.

## Plusieurs sociétés : un seul secret à gérer

Un jeton Pennylane est lié à une seule société. Pour ne pas avoir à
maintenir un secret par société, **toutes les clés vivent dans un seul
secret**, sous forme de liste JSON :

```json
["cle_api_societe_1", "cle_api_societe_2", "cle_api_societe_3"]
```

Le script boucle sur cette liste et identifie chaque société automatiquement
via l'API (endpoint `/me`) — aucun nom à saisir, aucune correspondance à
tenir à jour. **Ajouter une société = ajouter sa clé dans ce secret, relancer
le workflow. Rien d'autre à modifier.**

## Mise en route

### 1. Créer le dépôt GitHub

Créez un dépôt **privé** sur [github.com/new](https://github.com/new) (ne
cochez aucune case à la création, le dépôt doit rester vide).

### 2. Uploader les fichiers du projet

Sur la page du dépôt vide, cliquez sur le lien bleu "uploading an existing
file", puis glissez tous les fichiers et dossiers **visibles** du projet
décompressé.

⚠️ Le dossier `.github` (avec un point devant) est souvent ignoré par le
glisser-déposer, même rendu visible dans l'explorateur de fichiers. Si
`.github/workflows/report.yml` n'apparaît pas dans la liste des fichiers
uploadés : bouton **Add file → Create new file**, tapez le chemin complet
`.github/workflows/report.yml` dans le champ nom (GitHub crée le dossier
automatiquement), collez le contenu du fichier, Commit changes.

### 3. Déclarer les clés API

**Settings → Secrets and variables → Actions → New repository secret**
- Nom : `PENNYLANE_API_KEYS`
- Valeur : la liste JSON de vos clés, ex. `["cle_societe_1","cle_societe_2"]`
  (une seule société ? `["cle_societe_1"]` fonctionne aussi)

Chaque clé doit avoir, au minimum, les scopes en lecture sur *Trial Balance*
et *Fiscal Years* (Pennylane → Paramètres → Connectivité → Developers, au
moment de générer le jeton).

### 4. Premier lancement

**Actions → Export des balances Pennylane → Run workflow.**

## Récupérer le rapport

Le classeur est commité dans le dépôt à chaque exécution :
`reports/balances_dernier.xlsx` (toujours la dernière version) et
`reports/balances_YYYY-MM.xlsx` (une copie datée à chaque exécution, pour
garder un historique).

Pour le télécharger : ouvrez le fichier dans GitHub → bouton **⋯** (ou
"Download raw file") en haut à droite de l'aperçu → il se télécharge sur
votre ordinateur. Comme le dépôt est privé, seuls les comptes avec accès au
dépôt peuvent le voir.

Le workflow tourne automatiquement le 5 de chaque mois.

## Ajouter ou retirer une société plus tard

**Settings → Secrets and variables → Actions → `PENNYLANE_API_KEYS` →
Update.** Collez la nouvelle liste JSON complète (avec la clé ajoutée ou
retirée), Update secret, puis relancez le workflow manuellement une fois.

## Utilisation en local

```bash
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt

export PENNYLANE_API_KEYS='["cle_societe_1","cle_societe_2"]'
python export_balances.py
```

Résultat dans `reports/balances_dernier.xlsx`.

## Limites connues / points à vérifier

- **"Solde" = mouvement du mois**, pas un solde cumulé depuis le début de
  l'exercice. Pour un vrai bilan à une date donnée, il faudrait cumuler
  depuis l'ouverture — volontairement pas fait ici, à affiner si besoin.
- Les comptes sans aucun mouvement sur un mois donné n'apparaissent
  probablement pas dans la réponse `trial_balance` pour ce mois.
- Si une clé API échoue (jeton expiré, scope manquant...), le script
  continue avec les autres sociétés et signale l'échec en fin d'exécution
  plutôt que de tout arrêter.
- Le nombre d'appels API croît avec le nombre de mois disponibles × le
  nombre de sociétés (un appel par mois). Pour une société avec plusieurs
  années d'historique, la première exécution peut prendre quelques minutes.
- GitHub Actions gratuit : 2 000 minutes/mois sur dépôt privé, largement
  suffisant pour un usage mensuel.

## Documentation API utilisée

- [Balance comptable (trial_balance)](https://pennylane.readme.io/reference/gettrialbalance)
- [Exercices fiscaux (fiscal_years)](https://pennylane.readme.io/reference/company-fiscal-years)
- [Profil utilisateur / société (/me)](https://pennylane.readme.io/changelog/v2-new-user-profile-endpoint) —
  utilisé uniquement pour identifier automatiquement le nom de chaque société
