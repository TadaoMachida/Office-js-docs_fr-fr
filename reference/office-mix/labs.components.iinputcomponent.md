
# Labs.Components.IInputComponent

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Permet des interactions avec un composant de saisie.

```
interface IInputComponent extends Labs.Core.IComponent
```


## Propriétés


|Nom|Description|
|:-----|:-----|
| `maxScore: number`|Note maximale autorisée pour le composant de saisie.|
| `timeLimit: number`|Délai imparti à la résolution du problème du composant de saisie.|
| `hasAnswer: boolean`|Indique **True** si le composant a une réponse.|
| `answer: any`|Réponse au problème du composant, le cas échéant.|
| `secure: boolean`|Indique **True** si le composant de saisie est sécurisé.|
