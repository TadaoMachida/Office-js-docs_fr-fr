
# Propriété Context.touchEnabled
Obtient des informations indiquant si le complément est exécuté dans une application hôte Office tactile.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Dernière modification dans **|1.1|

```
var isTouchEnabled = Office.context.touchEnabled;
```


## Valeur renvoyée

Renvoie **True** si le complément est exécuté sur un appareil tactile, tel qu’un iPad ; sinon, renvoie **False**.


## Remarques

Utilisez la propriété **touchEnabled** pour déterminer quand votre complément est exécuté sur un appareil tactile et, si nécessaire, régler le type de contrôle, la taille et l’espacement des éléments dans l’interface utilisateur de votre complément afin d’adapter chaque interaction tactile.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office pour iPad**|
|:-----|:-----|
|**Excel**|v|
|**PowerPoint**|v|
|**Word**|v|

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Introduites|
