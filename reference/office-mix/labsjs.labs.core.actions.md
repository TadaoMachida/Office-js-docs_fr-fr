
# LabsJS.Labs.Core.Actions
Fournit une vue d’ensemble de l’API JavaScript LabJS.Labs.Core.Actions.

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Ces API représentent les opérations d’un atelier, en indiquant ses comportements en cours. Celles-ci vous permettent de créer des composants ou de développer des connexions avec un nouveau lecteur (autre qu’Office Mix).

## Module d’API LabsJS.Labs.Core.Actions

Le module Actions contient les types suivants :


### Interfaces


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|Composant à fermer.|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|Composant associé à la tentative.|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|Résultat de la création d’une tentative pour le composant donné.|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|Crée un composant.|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|Résultat [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) de la création d’un composant.|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|Résultat de l’action Obtention d’une valeur.|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|Résultat de l’envoi d’une réponse pour une tentative.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|Options disponibles pour l’expiration de la tentative en cours.|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|Options disponibles pour l’opération Obtention d’une valeur.|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|Options associées à une tentative de reprise.|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|Options disponibles pour l’action Envoi de la réponse.|

### Variables


|||
|:-----|:-----|
| `var CloseComponentAction: string`|Ferme le composant et indique qu’il ne fera plus l’objet d’actions.|
| `var CreateAttemptAction: string`|Action pour créer une tentative.|
| `var CreateComponentAction: string`|Action pour créer un composant.|
| `var AttemptTimeoutAction: string`|Action Expiration de la tentative.|
| `var GetValueAction: string`|Action pour récupérer une valeur associée à une tentative.|
| `var ResumeAttemptAction: string`|Action Reprise de la tentative. Indique que l’utilisateur relance une tentative donnée.|
| `var SubmitAnswerAction: string`|Action pour envoyer une réponse pour une tentative donnée.|
