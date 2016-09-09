
# Énumération SelectionMode
Spécifie s’il faut sélectionner (mettre en surbrillance) l’emplacement à atteindre (lorsque la méthode [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) est utilisée).

|||
|:-----|:-----|
|**Ajouté dans la version Office.js**|1.1|

|||
|:-----|:-----|
|**Hôtes :**|Excel, PowerPoint, Word|
|**Ajouté dans**|1.1|



```
Office.SelectionMode
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|L’emplacement sera sélectionné (mise en surbrillance).|
|Office.SelectionMode.None|"none"|Le curseur est déplacé au début de l’emplacement.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Introduit|
