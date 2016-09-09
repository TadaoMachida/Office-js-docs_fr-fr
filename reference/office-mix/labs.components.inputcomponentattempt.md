
# Labs.Components.InputComponentAttempt

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Tentative d’interaction avec un composant de saisie.

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## Méthodes




### constructeur

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crée une instance de la classe **InputComponentAttempt**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _labs_|Ateliers ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) associés à la tentative.|
| _componentID_|ID du composant associé à la tentative.|
| _attemptId_|ID de la tentative spécifique.|
| _values_|Tableau contenant les instances de valeur ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)).|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Passe en revue les actions récupérées pour la tentative spécifiée et renseigne l’état de l’atelier.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _action_|Action associée à l’état de l’atelier.|

### getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

Récupère tous les envois précédemment effectués pour la tentative spécifiée.


### submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

Envoie une nouvelle réponse notée par l’atelier. N’utilise pas l’hôte pour calculer une note.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _answer_|Réponse associée à la tentative.|
| _result_|Résultat associé à l’envoi.|
| _callback_|Fonction de rappel qui se déclenche une fois l’envoi reçu.|
