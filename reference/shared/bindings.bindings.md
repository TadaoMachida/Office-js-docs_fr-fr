
# Bindings, objet
Représente les liaisons du complément au sein du document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification** dans|1.1|

```js
Office.context.document.bindings
```


**Propriétés**

|||
|:-----|:-----|
|Nom|Description|
|[document](../../reference/shared/bindings.document.md)|Obtient un objet **Document** qui représente le document associé à cet ensemble de liaisons.|

**Méthodes**

|||
|:-----|:-----|
|Nom|Description|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|Ajoute une liaison à un élément nommé dans le document.|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|Affiche l’interface utilisateur qui permet à l’utilisateur de spécifier une sélection à lier.|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|Ajoute un objet de liaison du type spécifié lié à la sélection actuelle dans le document.|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|Obtient toutes les liaisons qui ont été créées précédemment.|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|Obtient la liaison spécifiée par son identificateur.|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|Supprime la liaison spécifiée.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office pour Bureau Windows|Office Online (dans un navigateur)|Office pour iPad|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad|
|1.1|Pour [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) et [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md), la liaison aux données de matrice en tant que liaison de tableau dans les compléments est désormais prise en charge pour Excel.|
|1.1|<ul><li>Pour la propriété <a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">document</a>, l’accès à un objet <span class="keyword">Document</span> qui représente la base de données Access actuelle dans des compléments du contenu pour Access est désormais possible.</li><li>Pour toutes les méthodes, la liaison de tableau dans les compléments de contenu pour Access est désormais prise en charge. </li></ul>|
|1,0|Introduit|
