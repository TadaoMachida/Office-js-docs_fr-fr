
# ProjectProjectFields, énumération
Spécifie les champs de projet disponibles en tant que paramètres pour la méthode **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Ajouté dans**|1,0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## Membres


****


|**Membre	**|**Description**|
|:-----|:-----|
|**CurrencyDigits**|Nombre de chiffres après la décimale pour la devise.|
|**CurrencySymbol**|Symbole de la devise.|
|**CurrencySymbolPosition**|Placement du symbole de la devise : Non spécifié = -1 ; Devant la valeur sans espace ($0) = 0 ; Derrière la valeur sans espace (0$) = 1 ; Devant la valeur avec un espace ($ 0) = 2 ; Derrière la valeur avec un espace (0 $) = 3.|
|**GUID**|GUID du projet.|
|**Finish**|Date de fin du projet.|
|**Démarrer**|Date de début du projet.|
|**ReadOnly**|Spécifie si le projet est en lecture seule.|
|**VERSION**|Version du projet.|
|**WorkUnits**|Unités de travail du projet, par exemple des jours ou des heures.|
|**ProjectServerUrl**|URL Project Web App, pour les projets stockés dans Project Server.|
|**WSSUrl**|L’URL SharePoint, pour les projets synchronisés avec une liste SharePoint.|
|**WSSList**|Nom de la liste SharePoint pour les projets synchronisés avec une liste de tâches.|

## Remarques

Une constante **ProjectProjectFields** peut être utilisée en tant que paramètre de la méthode **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|
|:-----|:-----|:-----|
|**Projet**|v||

|||
|:-----|:-----|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Méthode getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)
