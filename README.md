# Copier pour Odoo — complément Outlook

Ajoute un bouton **« Copier pour Odoo »** quand tu lis un mail (reçu ou envoyé) dans le nouvel Outlook ou Outlook web.
Un clic copie dans le presse-papier :

```
________________________________________
De: Jean Dupont <support@fournisseur.com>
Envoyé: Vendredi 30 octobre 2026 16:56
À: Marie Martin <m.martin@exemple.com>
Cc: support <support@exemple.com>
Objet: RE: Demande de support (#12345)

Bonjour Marie, ...
```

Il suffit ensuite de faire **Ctrl+V** dans la note Odoo (mise en forme conservée).

Options dans le volet :
- **Dernier message uniquement** (coché par défaut) : retire l'historique cité sous le message.
- **Retirer la signature** (coché par défaut) : garde la formule de politesse et le nom, retire titre, téléphone, adresse, logos, bannières.
- **Retirer les images** (coché par défaut) : enlève les images restantes (captures, logos).
- **Ligne de séparation** : la ligne `______` au-dessus de l'entête.

Les cases cochées sont mémorisées d'une ouverture à l'autre.

Le complément ne fait que **lire** le mail sélectionné (permission `ReadItem`). Rien n'est envoyé ailleurs.

---

## Installation (une seule fois, ~10 min)

Un complément Outlook doit être servi en HTTPS. Il faut donc d'abord publier les fichiers quelque part, puis déclarer le complément dans Outlook.

### 1. Héberger les fichiers

Fichiers à publier (tels quels, même arborescence) :

```
taskpane.html
format.js
assets/icon-16.png, icon-32.png, icon-64.png, icon-80.png, icon-128.png
```

Option la plus simple : **GitHub Pages** (gratuit)
1. Crée un dépôt, par ex. `copier-odoo`, et dépose-y ces fichiers.
2. *Settings → Pages → Branch : main / root → Save*.
3. L'adresse sera du type `https://<ton-compte>.github.io/copier-odoo`.

Tout autre hébergement HTTPS de pages statiques convient (serveur web Connectic, Azure Static Web Apps…). Une bibliothèque SharePoint ne convient pas.

### 2. Adapter le manifest

Ouvre `manifest.xml` dans un éditeur de texte et remplace **toutes** les occurrences de
`https://VOTRE-HEBERGEMENT` par ton adresse (sans `/` final), par ex. `https://torapwnz.github.io/copier-odoo`.

Vérifie que `https://.../taskpane.html` s'ouvre dans ton navigateur.

### 3. Ajouter le complément dans Outlook

1. Ouvre **https://aka.ms/olksideload** (ou dans Outlook : *Applications / Obtenir des compléments → Mes compléments*).
2. En bas : **Ajouter un complément personnalisé → Ajouter à partir d'un fichier…**
3. Sélectionne `manifest.xml` et confirme.

Le bouton apparaît alors dans le ruban (ou dans le menu **Applications** / « … ») quand un mail est ouvert.
Épingle le volet (icône punaise) pour qu'il reste ouvert quand tu passes d'un mail à l'autre.

> Si l'option « Ajouter un complément personnalisé » n'apparaît pas, l'ajout manuel est bloqué par la politique Microsoft 365 de Connectic.
> Il faut alors demander à l'admin M365 de le déployer : *Centre d'administration Microsoft 365 → Paramètres → Applications intégrées → Charger des applications personnalisées → manifest.xml* (possible pour toi seul ou pour toute l'équipe support).

---

## Utilisation au quotidien

1. Sélectionne le mail (dans *Éléments envoyés* pour tes envois, dans la boîte de réception pour les réponses).
2. Clique sur **Copier pour Odoo** → **Copier le mail**.
3. Dans la tâche Odoo : *Log note* → **Ctrl+V** → *Log*.

Si la copie automatique est bloquée, l'aperçu est sélectionné automatiquement : fais **Ctrl+C** puis colle dans Odoo.
