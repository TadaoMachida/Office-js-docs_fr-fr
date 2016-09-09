# Hosts, élément

Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres. 

Lorsqu’il est inclus dans le nœud [VersionOverrides](./versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste. 

## Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Hôte](#hôte)    |  Oui   |  Décrit un hôte et ses paramètres. |

> ** Remarque : ** Outlook doit `Hosts`contenir une définition `Host` pour `MailHost`.

---- 

## Élément Host
Indique un type particulier d’application Office où le complément doit être activé, comme « document », « classeur », « présentation », « projet », « boîte aux lettres » et « bloc-notes ».

### Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Décrit l’hôte d’Office auquel ces paramètres s’appliquent.|

### Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Oui   |  Définit le facteur de forme affecté. |


### xsi:type
Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’appliquent également les paramètres contenus. La valeur doit être l’une des suivantes :

- `MailHost` (Outlook)    


## Exemple de Hosts 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
