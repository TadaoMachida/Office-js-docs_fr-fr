# Objet settings.settingschangedeventargs
Fournit des informations sur les paramètres qui ont déclenché l’événement [settingsChanged](settings.settingschangedevent.md).

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel |
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans **|1,0|

```js
Office.EventType.SettingsChanged
```

## Membres

**Propriétés**

|**Nom**|**Description**|
|:-----|:-----|
|**[paramètres](settings.settingschangedeventargs.setting.md)**|Obtient un objet **Settings** qui représente les paramètres qui ont déclenché l’événement settingsChanged.|
|**[type](settings.settingschangedeventargs.type.md)**|Obtient une valeur d’énumération **EventType** qui identifie le genre d’événement déclenché.|

## Remarques

Pour ajouter un gestionnaire d’événements à l’événement **settingsChanged**, utilisez la méthode [addHandlerAsync](settings.addhandlerasync.md) de l’objet **Settings**.

L’événement **settingsChanged** se déclenche seulement lorsque le script de votre complément appelle la méthode **Settings.saveAsync** pour rendre persistante la copie en mémoire des paramètres dans le fichier de document. L’événement **settingsChanged** ne se déclenche pas lors de l’appel de la méthode [Settings.set](settings.set.md) ou [Settings.remove](settings.remove.md).

L’événement **settingsChanged** a été conçu pour vous permettre de gérer des conflits potentiels quand un ou plusieurs utilisateur(s) tente(nt) d’enregistrer des paramètres simultanément lorsque votre complément est utilisé dans un document partagé (co-créé).


 >**Important** : le code de votre complément peut inscrire un gestionnaire pour l’événement **settingsChanged** même lorsque le complément est exécuté avec un client Excel, mais l’événement ne se déclenche que si le complément est chargé avec une feuille de calcul ouverte dans Excel Online _et_ que plusieurs utilisateurs se servent de la feuille de calcul (co-création). Par conséquent, l’événement **settingsChanged** n’est réellement pris en charge que dans des scénarios de co-création Excel Online.



## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||


|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Paramètres|
|**Niveau d’autorisation minimal**|Restricted|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge

|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|
