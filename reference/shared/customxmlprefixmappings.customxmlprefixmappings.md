
# CustomXmlPrefixMappings, objet
Représente une collection de mappages de préfixes d’espace de noms personnalisés.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans **|1.1|

```
CustomXmlPrefixMappings
```


## Membres


**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|Ajoute de manière asynchrone un préfixe au mappage d’espace de noms à utiliser lors de l’interrogation d’un élément.|
|[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|Obtient de manière asynchrone l’espace de noms mappé au préfixe spécifié.|
|[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|Obtient de manière asynchrone le préfixe de l’espace de noms spécifié.|

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
