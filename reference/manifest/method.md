
# Élément Method
Spécifie une méthode individuelle de l’API JavaScript pour Office requise pour l’activation de votre complément Office.

 **Type de complément :** Application de contenu et de volet Office


## Syntaxe :


```XML
<Method Name="string "/>
```


## Contenu dans :

 _ [Méthodes](../../reference/manifest/methods.md)_


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.|

## Remarques

Les éléments **Methods** et **Method** ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de spécifications, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro).


 >**Important**  Étant donné qu’il n’existe aucun moyen de spécifier la version minimale requise pour les différentes méthodes, afin de vous assurer qu’une méthode est disponible lors de l’exécution, vous devez également utiliser une instruction **if** lorsque vous appelez cette méthode dans le script de votre complément. Pour plus d’informations sur la procédure à suivre, consultez l’article décrivant l’[API JavaScript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements).

