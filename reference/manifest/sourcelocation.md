
# Élément SourceLocation
Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<SourceLocation DefaultValue="string " />
```


## Contenu dans :

[DefaultSettings](../../reference/manifest/defaultsettings.md) (compléments de contenu et de volet Office)

[FormSettings](../../reference/manifest/formsettings.md) (compléments de messagerie)


## Peut contenir :

[Remplacer](../../reference/manifest/override.md)


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](../../reference/manifest/defaultlocale.md).|
