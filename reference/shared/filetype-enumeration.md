
# FileType, énumération
Spécifie le format de retour du document.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Dernière modification dans **|1.1|

```js
Office.FileType
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Retourne l’intégralité du document (.pptx ou .docx) au format Office Open XML (OOXML) sous forme de tableau d’octets.|
|Office.FileType.Pdf|"pdf"|Retourne l’intégralité du document au format PDF sous la forme d’un tableau d’octets.|
|Office.FileType.Text|"text"|Renvoie uniquement le texte du document sous forme d’une **chaîne**. (Word uniquement)|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire pour l’enregistrement au format PDF.|
|1.0|Introduit|
