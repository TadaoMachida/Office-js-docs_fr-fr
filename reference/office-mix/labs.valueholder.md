
# Labs.ValueHolder

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Objet conteneur comportant et suivant des valeurs pour un atelier spécifié. Les valeurs peuvent être stockées localement ou sur le serveur.

```
class ValueHolder<T>
```


## Variables


|||
|:-----|:-----|
| `public var isHint: boolean`|Indique **True** si la valeur est un conseil.|
| `public var hasBeenRequested: boolean`|Indique **True** si la valeur a été demandée par l’atelier.|
| `public var hasValue: boolean`|Indique **True** si le conteneur de la valeur a la valeur souhaitée.|
| `public var value: T`|Valeur conservée dans le conteneur.|
| `public var id: string`|ID de la valeur.|

## Méthodes




### getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

Récupère la valeur spécifiée.

 **Paramètres**


|||
|:-----|:-----|
| _callback_|Fonction de rappel qui renvoie la valeur spécifiée.|

### provideValue

 `public function provideValue(value: T): void`

Méthode interne qui fournit la valeur au conteneur de valeur.

 **Paramètres**


|||
|:-----|:-----|
| _value_|Valeur à fournir au conteneur de valeur.|
