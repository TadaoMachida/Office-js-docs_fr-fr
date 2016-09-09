
# Élément Sets
Spécifie le sous-ensemble minimal de l’API JavaScript pour Office nécessaire à l’activation de votre complément Office.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## Contenu dans :

[Configuration requise](../../reference/manifest/requirements.md)


## Peut contenir :

[Set](../../reference/manifest/set.md)


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|chaîne|facultatif|Spécifie la valeur de l’attribut **MinVersion** par défaut pour tous les éléments [Set](../../reference/manifest/set.md) enfants. La valeur par défaut est « 1.1 ».|

## Remarques

Pour plus d’informations sur les ensembles de conditions requises, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

