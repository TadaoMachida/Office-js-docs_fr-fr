
# Énumération ActiveView
Spécifie l’état de l’affichage dynamique du document, par exemple, si l’utilisateur peut modifier le document.

|||
|:-----|:-----|
|**Ajouté dans la version Office.js**|1.1|

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint|
|**Ajouté dans**|1.1|



```
Office.ActiveView
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.ActiveView.Read|« lecture »|L’affichage actif de l’application hôte permet seulement à l’utilisateur de lire le contenu du document.|
|Office.ActiveView.Edit|« édition »|L’affichage actif de l’application hôte permet à l’utilisateur de modifier le contenu du document.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint dans Office pour iPad.|
|1.1|Introduit|
