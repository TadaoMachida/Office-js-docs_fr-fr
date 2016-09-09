
# NodeDeletedEventArgs, objet
Fournit des informations sur le nœud supprimé qui a déclenché l’événement [dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md).

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Ajouté dans**|1.1|

```
NodeDeletedEventArgs
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Obtient des informations indiquant si le nœud a été supprimé dans le cadre d’une action Annuler/Rétablir effectuée par l’utilisateur.|
|[oldNextSibling](../../reference/shared/customxmlpart.oldnextsibling.md)|Obtient l’ancien frère suivant du nœud qui vient d’être supprimé de l’objet **CustomXMLPart**.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Obtient le nœud qui vient d’être supprimé de l’objet **CustomXmlPart**.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

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
