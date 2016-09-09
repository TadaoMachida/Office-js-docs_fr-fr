

# Événement

L’objet `event` est transmis en tant que paramètre aux fonctions de complément appelées par des boutons de commande sans interface utilisateur. Cet objet permet au complément d’identifier le bouton sur lequel l’utilisateur a cliqué et d’informer l’hôte que son traitement est terminé.

Par exemple, un bouton est défini dans un manifeste de complément de la manière suivante :

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

L’attribut `id` du bouton a pour valeur `eventTestButton`. Le bouton appelle la fonction `testEventObject` définie dans le complément. Cette fonction ressemble à ceci :

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

### Membres

####  source:Object

Obtient l’identificateur du bouton de commande du complément qui a appelé la méthode.

La propriété `source` renvoie un objet avec les propriétés suivantes.

| Propriété | Description |
| --- | --- |
| `id` | Valeur de l’attribut `id` de l’élément `Control` qui définit le bouton de commande du complément dans le manifeste de complément. |

Cette valeur peut être utilisée quand plusieurs boutons appellent la même fonction, mais vous devez effectuer différentes actions en fonction du bouton sur lequel l’utilisateur a cliqué.

##### Type :

*   Objet

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### Méthodes

####  completed()

Indique que le complément a terminé le traitement déclenché par le bouton de commande d’un complément.

Cette méthode doit être appelée à la fin d’une fonction qui a été appelée par une commande de complément définie avec un élément `Action` avec un attribut `xsi:type` ayant la valeur `ExecuteFunction`. Appeler cette méthode indique au client hôte que la fonction est terminée et qu’il peut nettoyer les états figurant dans l’appel de la fonction. Par exemple, si l’utilisateur ferme Outlook avant l’appel de cette méthode, Outlook vous avertit qu’une fonction est en cours d’exécution.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```