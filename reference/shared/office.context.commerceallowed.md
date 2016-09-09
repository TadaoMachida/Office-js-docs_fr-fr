
# Propriété Context.commerceAllowed
Obtient des informations indiquant si le complément est exécuté sur une plateforme qui autorise les liens vers des systèmes de paiement externes.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Dernière modification dans **|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## Valeur renvoyée

Renvoie **True** si les développeurs peuvent afficher l’interface utilisateur de vente ou de mise à niveau dans le complément sur cette plateforme ; sinon, renvoie **False**.


## Remarques

L’App Store iOS ne prend pas en charge les applications avec des compléments qui indiquent des liens vers d’autres systèmes de paiement. Toutefois, les compléments Office s’exécutant sur la version de bureau de Windows ou pour Office Online dans le navigateur autorisent les liens de ce type. Si vous voulez que l’interface utilisateur de votre complément indique un lien vers un système de paiement externe sur des plateformes autres qu’iOS, vous pouvez utiliser la propriété **commerceAllowed** pour contrôler l’affichage du lien.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour iPad**|
|:-----|:-----|
|**Excel**|v|
|**PowerPoint**||
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
