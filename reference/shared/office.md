

# Office, objet
Représente une instance d’un complément, qui permet d’accéder aux objets de niveau supérieur de l’API.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```js
Office
```


## Membres


**Propriétés**

|||
|:-----|:-----|
|Nom|Description|
|[context](../../reference/shared/office.context.md)|Obtient l’objet Context qui représente l’environnement d’exécution du complément et permet d’accéder aux objets de niveau supérieur de l’API.|
|[cast.item](../../reference/shared/office.cast.item.md)|Fournit la fonction IntelliSense dans Visual Studio pour les messages et rendez-vous en mode composition ou lecture. <br/><br/><blockquote>**Remarque**  Uniquement applicable au moment de la conception lorsque vous développez des compléments Outlook dans Visual Studio. </blockquote>|

**Méthodes**

|||
|:-----|:-----|
|Nom|Description|
|[select](../../reference/shared/office.select.md)|Crée une promesse de retour d’une liaison en fonction de la chaîne de sélecteur passée.|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|Active et désactive l’alias **Office** pour l’espace de noms **Microsoft.Office.WebExtension** complet.|

**Événements**

|||
|:-----|:-----|
|Nom|Description|
|[Initialiser](../../reference/shared/office.initialize.md)|Se produit quand l’environnement d’exécution est chargé et que le complément est prêt à interagir avec l’application et le document hébergé.|

## Remarques

L’objet **Office** permet au développeur d’implémenter une fonction de rappel pour l’événement Initialize et donne accès à l’objet [Context](../../reference/shared/context.md).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Projet**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Types de complément**|De contenu Outlook, du volet Office|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|<ul><li>Pour <a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a>, l’obtention du contexte d’exécution dans les compléments de contenu pour Access est désormais prise en charge.</p></li><li><p>Pour <a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a>, la sélection de liaisons de tableau dans les compléments de contenu pour Access est désormais prise en charge.</li><li>Pour <a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a>, les compléments de contenu pour Access sont désormais pris en charge.</li><li>Pour <a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a>, l’initialisation dans les compléments de contenu pour Access est désormais prise en charge.</li></ul>|
|1.0|Introduit|

