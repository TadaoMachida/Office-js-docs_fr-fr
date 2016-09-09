
# Binding, objet
Classe abstraite qui représente une liaison à une section du document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans les [ensembles de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**Dernière modification dans TableBinding**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## Membres


**Objets**


|**Nom**|**Description**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|Représente une liaison à deux dimensions de lignes et de colonnes.|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|Représente une liaison à deux dimensions de lignes et de colonnes, avec éventuellement des en-têtes.|
|[TextBinding](../../reference/shared/binding.textbinding.md)|Représente une sélection de texte lié dans le document.|

**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|Obtient l’objet **Document** associé à la liaison.|
|[id](../../reference/shared/binding.id.md)|Obtient l’identificateur de l’objet.|
|[type](../../reference/shared/binding.type.md)|Obtient le type de la liaison.|

**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|Ajoute un gestionnaire à la liaison pour le type d’événement spécifié.|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|Retourne les données contenues dans la liaison.|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|Supprime le gestionnaire spécifié de la liaison pour le type d’événement spécifié.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Écrit des données dans la section liée du document représenté par l’objet de liaison spécifié.|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Définit ou met à jour la mise en forme des éléments et données spécifiés dans le tableau lié.|

**Événements**


|**Nom**|**Description**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|Se produit quand des données sont modifiées dans la liaison.|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|Se produit quand la sélection est modifiée dans la liaison.|

## Remarques

L’objet **Binding** expose les fonctionnalités détenues par toutes les liaisons indépendamment du type.

L’objet **Binding** n’est jamais appelé directement. Il s’agit de la classe parente abstraite des objets qui représentent chaque type de liaison : [MatrixBinding](../../reference/shared/binding.matrixbinding.md), [TableBinding](../../reference/shared/binding.tablebinding.md) ou [TextBinding](../../reference/shared/binding.textbinding.md). Ces trois objets héritent des méthodes **getDataAsync** et **setDataAsync** de l’objet **Binding**, qui vous permettent d’interagir avec les données de la liaison. Elles héritent également des propriétés **id** et **type** pour l’interrogation des valeurs de propriétés correspondantes. En outre, les objets **MatrixBinding** et **TableBinding** exposent des méthodes supplémentaires pour les fonctionnalités relatives aux matrices et aux tableaux, par exemple le dénombrement des lignes et des colonnes.


## Informations de prise en charge


La prise en charge de chaque membre API de l’objet **Binding** diffère dans les applications hôtes Office. Voir la section « Informations de prise en charge » de la rubrique de chaque membre pour découvrir les informations de prise en charge d’hôte.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|MatrixBinding, TableBinding, TextBinding|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|
