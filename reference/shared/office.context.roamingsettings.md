
# Propriété Context.roamingSettings
Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément Outlook enregistré dans la boîte aux lettres d’un utilisateur.

|||
|:-----|:-----|
|**Hôtes :**|Outlook|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Boîte aux lettres|
|**Dernière modification dans **|1,0|

```
var appSettings = office.context.roamingSettings;
```


## Valeur renvoyée

Objet [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx).


## Remarques

L’objet **RoamingSettings** vous permet de stocker et d’accéder aux données pour un complément de messagerie conservé dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible pour le complément lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Boîte aux lettres|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|
