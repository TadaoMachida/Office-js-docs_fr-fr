
# Propriété Office.cast.item
Fournit la fonction IntelliSense pour les messages et rendez-vous en mode composition ou lecture.

|||
|:-----|:-----|
|**Hôtes :**|Outlook|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Boîte aux lettres|
|**Dernière modification dans **|1,0|



|||
|:-----|:-----|
|**Modes Outlook applicables**|Conception dans Visual Studio uniquement|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## Valeur renvoyée

Ensemble de méthodes permettant de sélectionner la fonction IntelliSense appropriée pour votre complément Outlook.


## Remarques

Cette propriété et ses méthodes prennent en charge IntelliSense pour le développement de complément Outlook uniquement sur Visual Studio. Elles n’ont pas d’effet sur d’autres outils de développement.

Les méthodes **Office.cast.item** sont utilisées au moment de la conception dans Visual Studio pour fournir la fonction IntelliSense spécifique pour la propriété **Office.context.mailbox.item**. Lorsque vous utilisez la méthode **toAppointmentCompose**, par exemple, IntelliSense n’affiche que les méthodes et propriétés **Appointment** qui s’appliquent au mode composition.

Lors de l’exécution, les méthodes **Office.cast.item** n’ont aucun effet sur votre complément Outlook.


## Exemple

L’exemple suivant utilise la méthode **toMessageCompose** pour effectuer une conversion de type de la propriété **Office.context.mailbox.item** afin de n’afficher la fonction IntelliSense que pour l’objet **Message** en mode composition. Une fois la conversion effectuée, la variable `message` n’affichera la fonction IntelliSense que pour les méthodes et propriétés qui peuvent être utilisées en mode composition.


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||Office pour Bureau Windows|Office Online (dans un navigateur)|Outlook pour Mac|
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



|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|
