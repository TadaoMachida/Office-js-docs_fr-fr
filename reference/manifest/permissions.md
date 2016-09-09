
# Élément Permissions
Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :

Pour les compléments du volet de tâches et de contenu :


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Pour les compléments de messagerie :




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## Contenu dans :

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Remarques

Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et celui décrivant les [autorisations de complément Outlook](../../docs/outlook/understanding-outlook-add-in-permissions.md).

