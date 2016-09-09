# Élément d’icône
Définit les éléments **Image** pour les contrôles de [bouton](./button.md) et de [menu](./menu-control.md).

## Éléments enfants
|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Image](#image)        | Oui |   Attribut resid d’une image à utiliser         |

## Image
Image du bouton. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image** dans l’élément **Images** dans l’élément [Resources](./resources.md). L’attribut **size** indique la taille de l’image en pixels. Trois tailles d’image sont requises (16, 32 et 80 pixels) et cinq autres tailles sont prises en charge (20, 24, 40, 48 et 64 pixels).|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  