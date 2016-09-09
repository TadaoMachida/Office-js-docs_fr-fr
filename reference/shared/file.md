
# File, objet
Représente le fichier du document associé à un complément Office.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Dernière modification dans **|1.1|

```
file
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|**[size](../../reference/shared/file.size.md)**|Obtient la taille du fichier de document en octets.|
|**[sliceCount](../../reference/shared/file.slicecount.md)**|Obtient le nombre de sections du fichier divisé.|

**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|**[closeAsync](../../reference/shared/file.closeasync.md)**|Ferme le fichier de document.|
|**[getSliceAsync](../../reference/shared/file.getsliceasync.md)**|Retourne la section spécifiée.|

## Remarques

Accédez à l’objet **File** avec la propriété [AsyncResult.value](../../reference/shared/asyncresult.value.md) dans la fonction de rappel transmise à la méthode [Document.getFileAsync](../../reference/shared/document.getfileasync.md).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office pour Bureau Windows|Office Online (dans un navigateur)|Office pour iPad|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans l’ensemble de ressources requis**|Fichier|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.0|Introduit|
