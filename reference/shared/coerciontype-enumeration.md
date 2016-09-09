
# CoercionType, énumération
Indique comment forcer le type des données retournées ou définies par la méthode appelée.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans la boîte aux lettres**|1.1|

```js
Office.CoercionType
```

## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|Renvoyer ou définir des données au format HTML.<br/><br/> **Remarque**  S’applique uniquement à des données dans des compléments pour Word et des compléments Outlook pour Outlook (mode composition).|
|Office.CoercionType.Matrix|"matrix"|Renvoyer ou définir des données sous forme de données tabulaires sans en-tête. Les données sont renvoyées ou définies sous forme d’un tableau de tableaux contenant des suites de caractères à une dimension. Par exemple, trois lignes de valeurs de type **string** dans deux colonnes correspondraient à ceci : ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`.<br/><br/> **Remarque**  S’applique uniquement aux données dans Excel et Word.|
|Office.CoercionType.Ooxml|"ooxml"|Renvoyer ou définir des données au format Office Open XML.<br/><br/> **Remarque**  S’applique uniquement aux données dans Word.|
|Office.CoercionType.SlideRange|"slideRange"|Renvoyer un objet JSON qui contient un tableau des ID, titres et index des diapositives sélectionnées. Par exemple, `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` pour une sélection de deux diapositives.<br/><br/> **Remarque**  S’applique uniquement aux données dans PowerPoint lors de l’appel de la méthode [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) pour obtenir la diapositive actuelle ou la plage sélectionnée de diapositives.|
|Office.CoercionType.Table|"table"|Renvoyer ou définir des données sous forme de données tabulaires avec en-têtes facultatifs. Les données sont renvoyées ou définies sous la forme d’un tableau de tableaux avec des en-têtes facultatifs.<br/><br/> **Remarque**  S’applique uniquement aux données dans Access, Excel et Word.|
|Office.CoercionType.Text|"text"|Renvoyer ou définir les données sous forme de texte (de type **string**). Les données sont renvoyées ou définies sous la forme d’une suite de caractères à une dimension.|
|Office.CoercionType.Image|"image"|Les données sont renvoyées ou définies sous la forme d’un flux d’images.<br/><br/> **Remarque**  S’applique uniquement aux données dans Excel, Word et PowerPoint.|
PowerPoint prend en charge uniquement **Office.CoercionType.Text**, **Office.CoercionType.Image** et **Office.CoercionType.SlideRange**.

Project prend en charge uniquement **Office.CoercionType.Text**.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Office pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|v|||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Projet**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Types de complément**|Contenu, Outlook (mode composition), volet Office|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.1|Prise en charge supplémentaire des [compléments Outlook en mode composition](../../docs/outlook/compose-scenario.md).|
|1.0|Introduit|
