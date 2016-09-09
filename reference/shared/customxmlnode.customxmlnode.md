
# CustomXmlNode, objet
Représente un nœud XML dans une arborescence au sein d’un document.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans **|1.1|

```js
CustomXmlNode
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|Obtient le nom de base du nœud sans le préfixe d’espace de noms, s’il en existe un.|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|Obtient le type de **CustomXMLNode**.|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|Récupère le GUID de chaîne de **CustomXMLPart**.|

**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|Obtient les nœuds de manière asynchrone sous la forme d’un tableau d’objets **CustomXMLNode** correspondant à l’expression XPath relative.|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|Obtient de manière asynchrone la valeur du nœud.|
|[getTextAsync](customxmlnode.gettextasync.md)|Obtient de manière asynchrone le texte d’un nœud XML dans une partie XML personnalisée.|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|Obtient de manière asynchrone le contenu XML du nœud.|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|Définit de manière asynchrone la valeur du nœud.|
|[setTextAsync](customxmlnode.settextasync.md)|Définit de manière asynchrone le texte d’un nœud XML dans une partie XML personnalisée.|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|Définit de manière asynchrone le contenu XML du nœud.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|CustomXmlParts|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word dans Office pour iPad.|
|1.0|Introduit|
