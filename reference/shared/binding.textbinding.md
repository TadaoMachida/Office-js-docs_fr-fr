
# TextBinding, objet
Représente une sélection de texte lié dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TextBindings|
|**Ajouté dans**|1,0|

```
TextBinding
```


## Remarques

L’objet **TextBinding** hérite de la propriété [id](../../reference/shared/binding.id.md), de la propriété [type](../../reference/shared/binding.type.md), de la méthode [getDataAsync](../../reference/shared/binding.getdataasync.md) et de la méthode [setDataAsync](../../reference/shared/binding.setdataasync.md) de l’objet [Binding](../../reference/shared/binding.md). Il n’implémente aucune propriété ou méthode supplémentaire.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|TextBindings|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.0|Introduit|
