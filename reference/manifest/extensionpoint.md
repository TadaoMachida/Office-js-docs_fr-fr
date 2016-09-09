# Élément ExtensionPoint

 Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office. L’élément **ExtensionPoint** est un élément enfant de [FormFactor](./formfactor.md). 

## Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Oui  | Type de point d’extension défini.|


## Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote

- **PrimaryCommandSurface** : ruban dans Office.
- **ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.

Les exemples suivants montrent comment utiliser l’élément  **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.


 >**Importante**  Pour les éléments qui contiennent un attribut ID, assurez-vous que vous indiquez un ID unique. Nous vous recommandons d’utiliser le nom de votre organisation, ainsi que votre ID. Par exemple, utilisez le format suivant.<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**Éléments enfants**
 
|**Élément**|**Description**|
|:-----|:-----|
|**CustomTab**|Obligatoire pour ajouter un onglet personnalisé au ruban (en utilisant  **PrimaryCommandSurface**). Si vous utilisez l’élément  **CustomTab**, vous ne pouvez pas utiliser l’élément  **OfficeTab**. L’attribut  **id** est requis.|
|**OfficeTab**|Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**. Pour plus d’informations, voir [OfficeTab](officetab.md).|
|**OfficeMenu**|Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : <br/> - **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. <br/> - **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.|
|**Group**|Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut comporter jusqu’à six contrôles. L’attribut  **id** est requis. Il s’agit d’une chaîne contenant un maximum de 125 caractères.|
|**Label**|Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Icône**|Obligatoire. Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. L’élément  **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément  **Resources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64.|
|**Tooltip**|Facultatif. Info-bulle du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément  **Resources**.|
|**Contrôle**|Chaque groupe exige au moins un contrôle. Un élément  **Control** peut être de type **Button** ou **Menu**. Utilisez  **Menu** pour spécifier une liste déroulante de contrôles de bouton. Actuellement, seuls les boutons et les menus sont pris en charge.Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](#contrôle-de-bouton) et [Contrôles de menu](#contrôles-de-menu).<br/>**Remarque**  Pour faciliter les opérations de dépannage, nous vous recommandons d’ajouter un élément **Control** et les éléments enfants **Resources** associés un par un.

## Points d’extension pour les commandes de complément Outlook

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](./formfactor.md).)

### CustomPane

Le point d’extension CustomPane définit un complément qui s’active lorsque des règles spécifiées sont respectées. Il est destiné uniquement au formulaire de lecture et s’affiche dans un volet horizontal. 

**Éléments enfants**

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **RequestedHeight** | Non |  Il s’agit de la hauteur demandée, en pixels, pour le volet d’informations lorsqu’il est exécuté sur un ordinateur de bureau. La taille peut être comprise entre 32 et 450 pixels.  |
|  **SourceLocation**  | Oui |  URL du fichier de code source du complément. Fait référence à un élément **Url** dans l’élément [Resources](./resources.md).  |
|  **Règle**  | Oui |  Règle ou ensemble de règles qui spécifie quand le complément doit être activé. Pour plus d’informations, voir [Règles d’activation pour les compléments Outlook](../../outlook/manifests/activation-rules.md). |
|  **DisableEntityHighlighting**  | Non |  Spécifie si la mise en surbrillance de l’entité doit être désactivée. |


#### Exemple CustomPane
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.

**Éléments enfants**

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](./customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### Exemple OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Exemple CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie. 

**Éléments enfants**

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](./customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### Exemple OfficeTab
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Exemple CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion. 

**Éléments enfants**

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](./customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### Exemple OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Exemple CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion. 

**Éléments enfants**

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](./customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

#### Exemple OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Exemple CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

Ce point d’extension place des boutons sur le ruban pour l’extension de module. 

**Éléments enfants**

|  Élément |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Ajoute les commandes à l’onglet de ruban par défaut.  |
|  [CustomTab](./customtab.md) |  Ajoute les commandes à l’onglet de ruban personnalisé.  |

