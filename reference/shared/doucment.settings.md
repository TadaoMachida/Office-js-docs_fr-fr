
# Settings, objet
Représente des paramètres personnalisés pour un complément de contenu ou de volet des tâches qui sont stockés dans le document hôte comme paires nom/valeur.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Dernière modification dans**|1.1|

```
Office.context.document.settings
```


## Membres


**Méthodes**

|||
|:-----|:-----|
|Nom|Description|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|Ajoute un gestionnaire d’événements pour l’événement **settingsChanged**.|
|[get](../../reference/shared/settings.get.md)|Récupère le paramètre spécifié.|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|Lit tous les paramètres persistants dans le document et actualise la copie du complément de ces paramètres en mémoire.|
|[remove](../../reference/shared/settings.remove.md)|Supprime le paramètre spécifié.|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|Supprime un gestionnaire d’événements pour l’événement **settingsChanged**.|
|[saveAsync](../../reference/shared/settings.saveasync.md)|Enregistre les paramètres.|
|[set](../../reference/shared/settings.set.md)|Définit ou crée le paramètre spécifié.|

**Événements**


|**Nom**|**Description**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|Se produit quand un paramètre est modifié.|

## Remarques

Les paramètres créés à l’aide des méthodes de l’objet **Settings** sont enregistrés pour chaque complément et pour chaque document. En d’autres termes, ils ne sont disponibles que pour le complément qui les a créés et uniquement dans le document où ils sont enregistrés.

Le nom d’un paramètre est une **chaîne**, alors que sa valeur peut être une **chaîne**, une donnée **numérique**, **booléenne**, **null**, un **objet** ou un **tableau**.

L’objet **Settings** est chargé automatiquement dans le cadre de l’objet [Document](../../reference/shared/document.md). En outre, il est disponible via l’appel de la propriété [settings](../../reference/shared/document.settings.md) de cet objet quand le complément est activé. Le développeur est responsable de l’appel de la méthode [saveAsync](../../reference/shared/settings.saveasync.md) après l’ajout ou la suppression de paramètres pour enregistrer les paramètres du document.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|v|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Paramètres|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Pour les méthodes <a href="7c4780cf-a779-4ac9-a362-c0bacae64a96.htm">addHandlerAsync</a> et <a href="735a255b-2a86-4b43-b1fa-e2a305815615.htm">removeHandlerAsync</a>, l’ajout et la suppression des gestionnaires d’événements pour l’événement <span class="keyword">SettingsChanged</span> dans les compléments de contenu pour Access sont désormais pris en charge. </p></li><li><p>Pour les méthodes <a href="aeac06dd-994e-4235-b208-1bd117395296.htm">get</a>, <a href="53a52c47-24b4-4d2d-b840-fe1b242cd795.htm">refreshAsync</a>, <a href="a92446bf-de65-45bd-8412-36ea8e77c5a2.htm">remove</a>, <a href="7147c221-937c-477c-98a6-f59d6200c27b.htm">saveAsync</a> et <a href="4e2c9758-953e-41e8-aca6-d8daf764a584.htm">set</a>, les paramètres personnalisés dans les compléments de contenu pour Access sont désormais pris en charge.</p></li></ul>|
|1.0|Introduit|

