
# Propriété NodeDeletedEventArgs.oldNode
Obtient le nœud qui vient d’être supprimé de l’objet **CustomXmlPart**.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans **|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## Valeur renvoyée

[CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) qui représente le nœud qui vient d’être supprimé.


## Remarques

Notez que ce nœud peut avoir des enfants, si une sous-arborescence est supprimée du document. En outre, ce nœud est un nœud « déconnecté » dont vous pouvez affiner l’interrogation vers le bas. Toutefois, vous ne pouvez pas effectuer l’interrogation vers le haut de l’arborescence (le nœud semble exister de manière isolée).


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




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word dans Office pour iPad.|
|1.0|Introduit|
