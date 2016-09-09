
# TableBinding, objet
Représente une liaison à deux dimensions de lignes et de colonnes, avec éventuellement des en-têtes.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Dernière modification dans la sélection**|1.1|

```
TableBinding
```


## Membres


**Propriétés**


|**Nom**|**Description**|**Mises à jour pour Office.js version 1.1**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|Obtient le nombre de colonnes de l’objet **TableBinding** spécifié.|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access.|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|Renvoie true si l’objet **TableBinding** spécifié comporte des en-têtes, sinon false.|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access.|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|Nombre de lignes de l’objet **TableBinding** spécifié.|Pour des raisons de performances, toujours renvoyer -1 dans les compléments de contenu pour Access.|

**Méthodes**


|**Nom**|**Description**|**Mises à jour pour Office.js version 1.1**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|Ajoute des colonnes et des valeurs à un tableau.||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|Ajoute des lignes et des valeurs à un tableau.|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access.|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|Efface la mise en forme du tableau lié.|Nouveauté dans Office.js version 1.1 pour les compléments pour Excel.|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|Supprime toutes les lignes et leurs valeurs (à l’exception des lignes d’en-tête) du tableau, en progressant de manière appropriée pour l’application hôte.|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Écrit des données dans la section liée du document représenté par l’objet de liaison spécifié.|<ul><li>Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access.</li><li>Prise en charge supplémentaire de la définition de la mise en forme lors de l’écriture de données dans des tableaux liés dans des compléments pour Excel.</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Définit la mise en forme de cellule et de tableau sur des éléments et des données spécifiés dans le tableau lié.|Peut définir la mise en forme de tableau dans les compléments pour Excel.|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|Met à jour les options de mise en forme de tableau sur le tableau lié.|Peut définir la mise en forme de tableau dans les compléments pour Excel.|

## Remarques

L’objet **TableBinding** hérite de la propriété [id](../../reference/shared/binding.id.md), de la propriété [type](../../reference/shared/binding.type.md), de la méthode [getDataAsync](../../reference/shared/binding.getdataasync.md) et de la méthode [setDataAsync](../../reference/shared/binding.setdataasync.md) de l’objet abstrait [Binding](../../reference/shared/binding.md).

Une fois que vous avez établi une liaison de tableau dans Excel, chaque nouvelle ligne ajoutée au tableau par un utilisateur est automatiquement incluse dans la liaison (**rowCount** augmente).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|TableBindings|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de [la définition de la mise en forme lors de l’insertion de tableaux](../../docs/excel/format-tables-in-add-ins-for-excel.md) dans Excel.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
