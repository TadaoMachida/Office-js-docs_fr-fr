
# HighResolutionIconUrl, élément
Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<HighResolutionIconUrl DefaultValue="string " />
```


## Peut contenir :

[Remplacer](../../reference/manifest/override.md)


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|chaîne (URL)|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](../../reference/manifest/defaultlocale.md).|

## Remarques

Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.

L’image doit être dans l’un des formats de fichier suivants, avec une résolution recommandée de 64 x 64 pixels : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application ou complément_ dans [Création d’applications et de compléments efficaces pour l’Office Store](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

