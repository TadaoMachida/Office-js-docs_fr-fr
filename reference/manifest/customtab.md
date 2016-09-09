# Élément CustomTab
Sur le ruban, indiquez l’onglet et le groupe où placer leurs commandes de complément. Il peut s’agir de l’onglet par défaut (soit  **Accueil**,  **Message**, ou  **Réunion**), ou d’un onglet personnalisé défini par le complément.

Sous les onglets personnalisés, le complément peut créer jusqu’à 10 groupes. Chaque groupe est limité à 6 contrôles, quel que soit l’onglet où il apparaît. Les compléments sont limités à un onglet personnalisé.

L’attribut **id** doit être unique au sein du manifeste.

## Éléments enfants
|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Group](./group.md)      | Oui |  Définit un groupe de commandes.  |
|  [Label](#label)      | Oui |  Étiquette pour CustomTab ou Group.  |
|  [Contrôle](#contrôle)    | Oui |  Ensemble d’un ou de plusieurs objets Control  |

## Group
Obligatoire. Voir [Élément group](./group.md).

## Label (Tab)
Obligatoire. Étiquette de l’onglet personnalisé. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément [ShortStrings](./resources.md#shortstrings) de l’élément [Resources](./resources.md).


##  Exemple CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```