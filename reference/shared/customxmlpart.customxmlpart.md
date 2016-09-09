
# CustomXmlPart, objet
Représente un **CustomXMLPart** unique dans une collection [CustomXMLParts](../../reference/shared/customxmlparts.customxmlparts.md).

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans **|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|Obtient une valeur qui indique si CustomXMLPart est prédéfini.|
|[id](../../reference/shared/customxmlpart.id.md)|Obtient le GUID de CustomXMLPart.|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|Obtient l’ensemble des mappages de préfixes d’espace de noms (CustomXMLPrefixMappings) utilisés pour le CustomXMLPart actuel.|

**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|Ajoute un gestionnaire d’événements de manière asynchrone pour un événement d’objet **CustomXmlPart**.|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|Supprime de manière asynchrone cette partie XML personnalisée de la collection.|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|Obtient de manière asynchrone les CustomXmlNodes de cette partie XML personnalisée qui correspondent au XPath spécifié.|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|Obtient de manière asynchrone le code XML contenu dans cette partie XML personnalisée.|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|Supprime un gestionnaire d’événements pour un événement d’objet **CustomXmlPart**.|

**Événements**


|**Nom**|**Description**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|Se produit quand un nœud est supprimé.|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|Se produit quand un nœud est inséré.|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|Se produit quand un nœud est remplacé.|

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
