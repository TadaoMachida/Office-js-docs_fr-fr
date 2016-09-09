
# Propriété officeTheme.controlForegroundColor
Obtient la couleur de premier plan du contrôle du thème Office.

 **Important :** Cette API fonctionne actuellement uniquement dans Excel, Outlook, PowerPoint et Word dans [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) sur le bureau Windows.



|||
|:-----|:-----|
|**Hôtes :**|Excel, Outlook, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Pas dans un ensemble|
|**Ajouté dans**|1.3|

```js
var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
```


## Valeur renvoyée

Tripler de couleur hexadécimal


## Remarques

Les couleurs renvoyées correspondent aux valeurs du thème Office, sélectionné par l’utilisateur en accédant à **Fichier**  >  **Compte Office**  >  **Thème Office**, qui est appliqué à toutes les applications hôtes Office.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|v||||
|**Outlook**|v||||
|**PowerPoint**|v||||
|**Word**|v||||



|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.3|Introduit|
