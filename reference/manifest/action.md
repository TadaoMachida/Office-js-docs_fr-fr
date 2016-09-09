# Action, élément
 Indique l’action à réaliser lorsque l’utilisateur sélectionne des contrôles de [bouton](./button-control.md) ou de [menu](./menu-control.md).
 
## Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Type d’action à effectuer|


## Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Spécifie le nom de la fonction à exécuter. |
|  [SourceLocation](#sourcelocation) |    Spécifie l’emplacement du fichier source pour cette action. |
  

## xsi:type
Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :
- ExecuteFunction
- ShowTaskpane

## FunctionName
Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](./functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## SourceLocation
Élément obligatoire lorsque  **xsi:type** est « ShowTaskpane ». Indique l’emplacement du fichier source pour cette action. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément [Urls](./resources.md#urls) dans l’élément [Resources](./resources.md).

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  
