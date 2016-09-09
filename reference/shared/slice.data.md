
# Propriété Slice.data
Obtient les données brutes de la section de fichier.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Dernière modification dans **|1.1|

```
var sliceData = slice.data;
```


## Valeur renvoyée

Données brutes de la section de fichier au format **Office.FileType.Text** ("text") ou **Office.FileType.Compressed** ("compressed"), tel que spécifié par le paramètre _fileType_ de l’appel à la méthode [Document.getFileAsync](../../reference/shared/document.getfileasync.md).


## Remarques

Les fichiers au format "compressed" retournent un tableau d’octets qui peut être transformé en chaîne encodée au format base64, si nécessaire.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Fichier|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.0|Introduit|
