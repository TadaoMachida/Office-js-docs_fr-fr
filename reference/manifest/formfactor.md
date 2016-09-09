# Élément FormFactor

Spécifie les paramètres d’un complément pour un facteur de forme donné. Par exemple, la définition d’un `Host` avec le type `MailHost` et `DesktopFormFactor` s’applique à Outlook pour le bureau, mais _pas_ à Outlook Web App ou à Outlook.com. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.

Chaque définition de facteur de forme contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](./functionfile.md) et [Élément ExtensionPoint](./extensionpoint.md). 

Les facteurs de forme suivants sont pris en charge :

- `DesktopFormFactor` (Office pour les clients Windows ou Mac)

## Éléments enfants

| Élément                               | Obligatoire | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](./functionfile.md)     | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [GetStarted](./getstarted.md)         | Non       | Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint. |

## Exemple FormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
