
# Élément Set
Spécifie un ensemble de conditions requises de l’API JavaScript pour Office nécessaires à l’activation de votre complément Office.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<Set Name="string " MinVersion="n .n ">
```


## Contenu dans :

[documents](../../reference/manifest/sets.md)


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Nom d’un [ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).|
|MinVersion|chaîne|facultatif|Spécifie la version minimale de l’ensemble d’API requis par votre complément. Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [Sets](../../reference/manifest/sets.md).|

## Remarques

Pour plus d’informations sur les ensembles de conditions requises, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-office-hosts-and-api-requirements).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).


 >**Important**  Pour les compléments de messagerie, il n’existe qu’un ensemble de conditions requises `"Mailbox"` disponible. Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier l’ensemble de conditions requises `"Mailbox"` dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office). De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.

