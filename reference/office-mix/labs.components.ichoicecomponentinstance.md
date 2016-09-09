
# Labs.Components.IChoiceComponentInstance

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Instance d’un composant de choix.

```
interface IChoiceComponentInstance extends Labs.Core.IComponentInstance
```


## Propriétés


|Nom|Description|
|:-----|:-----|
| `choices: Components.IChoice[]`|Tableau représentant la liste de choix associée au problème.|
| `timeLimit: number`|Délai imparti à la résolution du problème.|
| `maxAttempts: number`|Nombre maximal de tentatives autorisé pour le problème.|
| `maxScore: number`|Note maximale pour le problème.|
| `hasAnswer: boolean`|Indique **True** si le problème a une réponse.|
| `answer: any`|Réponse pour ce problème. Tableau si plusieurs réponses sont prises en charge, ou ID unique si une seule réponse est prise en charge.|
| `secure: boolean`|Indique si le questionnaire est sécurisé (les champs sécurisés sont cachés à l’utilisateur).|
