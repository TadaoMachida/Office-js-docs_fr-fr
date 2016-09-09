
# API de complément Outlook

Pour utiliser des API dans votre complément Outlook, vous devez spécifier l’emplacement de la bibliothèque Office.js, l’ensemble des conditions requises, le schéma et les autorisations.

## Bibliothèque Office.js

Pour interagir avec l’API du complément Outlook, vous devez utiliser les API JavaScript dans Office.js. Le CDN de la bibliothèque est _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js _. Les compléments envoyés à l’Office Store doivent faire référence à Office.js via ce CDN, mais ils ne peuvent pas utiliser de référence locale. 

Déclarez le CDN dans la balise **head** de la page web (fichier .html, .aspx ou .php) qui implémente l’interface utilisateur de votre complément, dans l’attribut **src** de la balise **script** :


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

L’ajout de nouvelles API ne modifie pas l’URL vers Office.js. La version de l’URL sera modifiée uniquement si un comportement d’API existant est interrompu.

> **Important :** Quand vous développez un complément pour une application hôte Office, référencez l’interface API JavaScript pour Office à partir de l’intérieur de la section `<head>` de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Au-delà de ce délai, un message d’erreur indiquant que le complément ne répond pas s’affiche à l’écran.  

## Ensembles de conditions requises

Toutes les API Outlook appartiennent à l’ensemble de conditions requises Mailbox. Celui-ci possède plusieurs versions et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. L’ensemble d’API le plus récent ne sera pas pris en charge par tous les clients Outlook à sa publication, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble seront également prises en charge. 

Pour savoir dans quels clients Outlook le complément s’affiche, indiquez la version de l’ensemble de conditions requises dans le manifeste. Par exemple, si vous indiquez la version 1.3 de l’ensemble de conditions requises, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge la version minimale 1.3. 

Le fait de spécifier un ensemble de conditions requises ne limite pas votre complément aux API de cette version. Si le complément spécifie l’ensemble de conditions requises version 1.1, mais est exécuté dans un client Outlook prenant en charge la version 1.3, le complément peut toujours utiliser les API v1.3. L’ensemble de conditions requises contrôle uniquement les clients Outlook dans lesquels le complément apparaît.

Pour vérifier la disponibilité des API à partir d’un ensemble de conditions requises de version supérieure à celle spécifiée dans le manifeste, vous pouvez utiliser l’API JavaScript standard :


```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **Remarque :** Ces vérifications ne sont pas nécessaires pour les API appartenant à l’ensemble de conditions requises dont la version est la même que celle spécifiée dans le manifeste.

Spécifiez l’ensemble de conditions requises minimal prenant en charge l’ensemble d’API critique pour votre scénario, sans lequel les fonctionnalités de votre complément ne fonctionneront pas. Vous pouvez indiquer l’ensemble de conditions requises dans le manifeste dans les éléments **Requirements**, **Sets** et **Set**. Pour plus d’informations, voir [Manifestes des compléments Outlook](../outlook/manifests/manifests.md) et [Présentation de l’ensemble de conditions requises pour les API Outlook](..\..\reference\outlook\tutorial-api-requirement-sets.md).

L’élément **Methods** ne s’applique pas aux compléments Outlook, ainsi vous ne pouvez pas déclarer la prise en charge de méthodes spécifiques.


## Autorisations

Votre complément requiert les autorisations appropriées pour utiliser les API dont il a besoin. Il existe quatre niveaux d’autorisations possibles. Pour plus de détails, voir [Présentation des autorisations du complément Outlook](../outlook/understanding-outlook-add-in-permissions.md).


|**Niveau d’autorisation**|**Description**|
|:-----|:-----|
|Restricted|Permet l’utilisation d’entités, mais pas d’expressions régulières.|
|Lire l’élément|En plus des autorisations indiquées dans _Restricted_, il autorise :<ul><li>expressions régulières</li><li>l’accès en lecture de l’API du complément Outlook</li><li>l’obtention des propriétés de l’élément et du jeton de rappel</li></ul>|
|Lecture/écriture|En plus des autorisations indiquées dans _Read item_, il autorise :<ul><li>l’accès total à l’API du complément Outlook, à l’exception de <b>makeEwsRequestAsync</b></li><li>la définition des propriétés de l’élément</li></ul>|
|Lire/écrire dans la boîte aux lettres|En plus des autorisations indiquées dans _Read/write_, il autorise :<ul><li>la création, la lecture, l’écriture d’éléments et de dossiers</li><li>l’envoi d’éléments</li><li>l’appel de [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext)</li></ul>|
En règle générale, vous devez spécifier l’autorisation minimale nécessaire à votre complément. Les autorisations sont déclarées dans l’élément **Permissions** du manifeste. Pour plus d’informations, voir [Manifestes des compléments Outlook](../outlook/manifests/manifests.md). Pour plus d’informations sur les questions de sécurité, voir [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/../../docs/develop/privacy-and-security.md).


## Ressources supplémentaires

- [Manifestes des compléments Outlook](../outlook/manifests/manifests.md)

- [Présentation de l’ensemble de conditions requises pour les API Outlook](../../reference/outlook/tutorial-api-requirement-sets.md)
    
- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
