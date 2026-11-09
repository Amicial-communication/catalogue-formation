# Catalogue de formation Amicial

Site : https://amicial-communication.github.io/catalogue-formation/

| Fichier | Rôle |
|---|---|
| `catalogue.xlsx` | Copie publiée du catalogue. **Ne pas la modifier ici** : le robot la remplace par la version SharePoint. |
| `index.html` | La page web qui lit l'Excel. |
| `logo.png` | Logo affiché en haut de page et sur le PDF. |
| `.github/` | Le robot de synchronisation et son contrôle. |

## Comment le catalogue se met à jour

1. La référente formation modifie `catalogue.xlsx` **dans le dossier réservé de l'intranet** (accès communication et RH uniquement) (formations, sessions, champs d'inscription, réglages).
2. Plusieurs fois par jour (toutes les 4 à 7 heures environ, au rythme choisi par GitHub), le robot télécharge ce fichier, le contrôle et le publie s'il a changé. Besoin d'aller plus vite : voir « Publier tout de suite ».
3. Une formation en cours de rédaction reste au statut **Brouillon** : elle n'apparaît pas tant qu'elle n'est pas « Publié ».

**Si le fichier pose problème** (lien expiré, onglet renommé, fichier illisible, formations publiées divisées par deux d'un coup), le robot ne publie rien : le site garde la version précédente et un ticket « Synchronisation du catalogue bloquée » s'ouvre dans l'onglet *Issues* (avec notification par e-mail). Il se ferme tout seul quand tout repart. Les points non bloquants (ID en double, date illisible…) sont listés dans le détail de chaque passage, onglet *Actions*.

## Réglage initial (une fois)

1. Sur l'intranet, créer un dossier réservé (Gérer l'accès → ne plus hériter des autorisations ; accès « Modifier » pour la communication et la référente RH et son suppléant). Y déposer la version actuelle de `catalogue.xlsx` : c'est elle qui fait foi désormais. La bibliothèque de l'équipe RH ne convient pas : le partage « Toute personne » y est bloqué.
2. Sur ce fichier dans SharePoint : **Partager → ⚙ Paramètres du lien → « Toute personne disposant du lien »**, **décocher « Autoriser la modification »**, pas de date d'expiration → **Copier le lien**. Le lien vise le fichier seul, pas le dossier.
3. Dans ce dépôt : **Settings → Secrets and variables → Actions → New repository secret**. Nom : `SHAREPOINT_URL`, valeur : le lien copié → **Add secret**.
4. Onglet **Actions → Synchronisation du catalogue → Run workflow**. Le passage doit finir en vert, avec « Catalogue publié » ou « Aucun changement » dans son résumé.

Le lien reste secret : il n'apparaît ni dans le code ni dans les journaux.

## Publier tout de suite

Onglet **Actions → Synchronisation du catalogue → Run workflow**, ou attendre le prochain passage (dans les heures qui suivent).

**Facultatif, publication en une minute** : un flux Power Automate « Lorsqu'un fichier est créé ou modifié (propriétés uniquement) » sur le dossier SharePoint, suivi de l'action GitHub « Create a repository dispatch event » (dépôt `Amicial-communication/catalogue-formation`, type d'événement `catalogue-modifie`). Ces deux connecteurs sont standard, sans licence Premium.

## Bon à savoir

- Modifier le fichier sur place. Le renommer ou le déplacer dans le même site garde le lien ; le supprimer puis en déposer un nouveau, ou le déplacer vers un autre site, casse le lien (le robot prévient) : il faut alors recréer le lien et mettre à jour le secret `SHAREPOINT_URL`.

- GitHub suspend les tâches programmées d'un dépôt public resté 60 jours sans activité. Le robot fait un « commit vide » tous les 45 jours sans changement pour l'éviter ; si jamais il est suspendu, GitHub prévient par e-mail et on le réactive en un clic dans l'onglet Actions.
- Le dépôt est public : le fichier publié est téléchargeable. N'y mettez rien de confidentiel (colonne « Relecture » comprise).
- Publication manuelle toujours possible : **Add file → Upload files** → `catalogue.xlsx` → **Commit**. Le robot la remplacera au passage suivant si le fichier de l'intranet est différent : mettez donc aussi l'intranet à jour.
