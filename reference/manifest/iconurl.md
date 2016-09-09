
# IconUrl, élément
Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<IconUrl DefaultValue="string " />
```


## Peut contenir :

[Remplacer](../../reference/manifest/override.md)


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|chaîne|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](../../reference/manifest/defaultlocale.md).|

## Remarques

Pour un complément de messagerie, l’icône s’affiche dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments** (Outlook) ou sous **Paramètres**  >  **Gérer les compléments** (Outlook Web App). Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer**  >  **Compléments**. Pour tous les types de compléments, l’icône est également utilisée sur le site de l’Office Store si vous publiez votre complément dans l’Office Store.

L’image doit être dans l’un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour les applications de volet Office et du contenu, l’image spécifiée doit faire 32 x 32 pixels. Pour les applications de messagerie, l’image doit faire 64 x 64 pixels. Vous devez également spécifier une icône à utiliser avec les applications hôtes Office en cours d’exécution sur des écrans haute résolution (DPI) à l’aide de l’élément [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md). Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application ou complément_ dans [Création d’applications et de compléments efficaces pour l’Office Store](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

