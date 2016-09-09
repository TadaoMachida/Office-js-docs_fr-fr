
# FilterType, énumération
Spécifie si le filtrage à partir de l’application hôte est appliqué quand les données sont récupérées.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Project, Word|
|**Dernière modification dans **|1.1|

```js
Office.FilterType
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.FilterType.All|« tout »|Retourne toutes les données (non filtrées par l’application hôte).|
|Office.FilterType.OnlyVisible|"onlyVisible"|Retourne uniquement les données visibles (filtrées par l’application hôte).|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.


Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Projet**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge

|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.0|Introduit|
