
# Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook

Vous pouvez indiquer des règles d’expressions régulières pour activer un complément Outlook dans certains scénarios de lecture. Lorsque l’utilisateur affiche un message ou un rendez-vous dans le volet de lecture ou l’inspecteur, Outlook évalue les règles d’expressions régulières dans le but de déterminer s’il doit activer votre complément. Ces règles ne sont pas évaluées par Outlook quand l’utilisateur compose un élément. Il existe également d’autres scénarios dans lesquels Outlook n’active pas les compléments ; par exemple, les éléments protégés par la Gestion des droits relatifs à l’information ou ceux présents dans le dossier Courrier indésirable. Pour plus d’informations, voir [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md).

Vous pouvez spécifier une expression régulière dans le cadre d’une règle [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) ou [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) dans le manifeste XML de complément. Outlook évalue les expressions régulières en fonction des règles définies pour l’interpréteur JavaScript utilisé par le navigateur de l’ordinateur client. Outlook prend en charge la même liste de caractères spéciaux que tous les processeurs XML. Le tableau suivant répertorie ces caractères spéciaux. Vous pouvez les utiliser dans une expression régulière en spécifiant la séquence d’échappement pour le caractère correspondant, comme décrit dans le tableau suivant.



|**Caractère**|**Description**|**Séquence d’échappement à utiliser**|
|:-----|:-----|:-----|
|"|Guillemets doubles|&amp;quot;|
|&amp;|Esperluette|&amp;amp;|
|'|Apostrophe|&amp;apos;|
|<|Signe inférieur à|&amp;lt;|
|>|Signe supérieur à|&amp;gt;|

## Règle ItemHasRegularExpressionMatch


La règle  **ItemHasRegularExpressionMatch** est très utile dans le contrôle de l’activation d’un complément basé sur les valeurs spécifiques d’une propriété prise en charge. La règle **ItemHasRegularExpressionMatch** contient les attributs ci-dessous.



|**Nom de l’attribut**|**Description**|
|:-----|:-----|
|**RegExName**|Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.|
|**RegExValue**|Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément doit être affiché.|
|**PropertyName**|Spécifie le nom de la propriété par rapport à laquelle l’expression régulière sera évaluée. Les valeurs autorisées sont  **BodyAsHTML**,  **BodyAsPlaintext**,  **SenderSMTPAddress** et **Subject**. Si vous spécifiez  **BodyAsHTML**, Outlook applique l’expression régulière uniquement si le corps de l’élément est de type HTML, sinon Outlook ne renvoie aucune correspondance pour cette expression régulière. Comme les rendez-vous sont toujours enregistrés au format RTF, une expression régulière qui spécifie  **BodyAsHTML** ne correspond à aucune chaîne dans le corps des éléments de rendez-vous.Si vous spécifiez  **BodyAsPlaintext**, Outlook applique toujours l’expression régulière au corps de l’élément.|
|**IgnoreCase**|Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par  **RegExName**.|

### Meilleures pratiques pour l’utilisation d’expressions régulières dans les règles

Prêtez une attention particulière aux éléments suivants lorsque vous utilisez des expressions régulières :


- Si vous spécifiez une règle  **ItemHasRegularExpressionMatch** dans le corps d’un élément, l’expression régulière doit filtrer également le corps et ne doit pas tenter de retourner la totalité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour essayer d’obtenir la totalité du corps d’un élément ne retourne pas toujours les résultats attendus.
    
- Le corps en texte brut renvoyé sur un navigateur peut être légèrement différent sur un autre. Si vous utilisez une règle [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) avec **BodyAsPlaintext** comme attribut **PropertyName**, testez votre expression régulière sur tous les navigateurs pris en charge par votre complément.
    
    Comme différents navigateurs utilisent diverses méthodes pour obtenir le corps du texte d’un élément sélectionné, vous devez vous assurer que votre expression régulière prend en charge les fines différences qui peuvent être renvoyées dans le cadre du corps de texte. Par exemple, certains navigateurs, comme Internet Explorer 9, utilisent la propriété  **innerText** du DOM, tandis que d’autres, comme Firefox, utilisent la méthode **.textContent()** afin d’obtenir le corps du texte d’un élément. En outre, différents navigateurs peuvent renvoyer des sauts de ligne de manière différente : un saut de ligne correspond à « \r\n » sur Internet Explorer et « \n » dans Firefox et Chrome. Pour plus d’informations, voir la rubrique sur la [compatibilité DOM W3C - HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07).
    
- Le corps HTML d’un élément est légèrement différent entre un client riche Outlook et Outlook Web App ou OWA pour périphériques. Définissez soigneusement vos expressions régulières. Prenons par exemple l’expression régulière suivante utilisée dans une règle  **ItemHasRegularExpressionMatch** avec **BodyAsHTML** comme valeur de l’attribut **PropertyName** :
    
```
      http.*\.contoso\.com
```


    A rule with this regular expression would match the string "http-equiv="Content-Type" which exists in the HTML body of an item in an Outlook rich client, as part of the following  **META** tag:
    

```HTML
      <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
```


La même règle ne renvoie pas cette correspondance dans Outlook Web App et OWA pour les appareils, car le corps HTML sur ces hôtes n’inclut pas la balise **META**. Cela peut avoir une influence sur le fait que le complément est activé correctement ou non pour les divers clients Outlook. Dans cet exemple, utilisez plutôt l’expression régulière suivante :
    

```
      http://.*\.contoso\.com/
```

- En fonction de l’application hôte, du type de périphérique ou de la propriété à laquelle l’expression régulière est appliquée, il existe d’autres meilleures pratiques et limites pour chaque hôte, que vous devez connaître lorsque vous créez des expressions régulières comme règle d’activation. Pour plus d’informations, voir [Limites d’activation et d’API JavaScript des compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    

### Exemples

La règle  **ItemHasRegularExpressionMatch** suivante active le complément chaque fois que l’adresse de messagerie SMTP de l’expéditeur correspond à « @contoso », indépendamment des caractères majuscules et minuscules.


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

L’exemple suivant montre une autre manière de spécifier la même expression régulière à l’aide de l’attribut  **IgnoreCase**.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

La règle  **ItemHasRegularExpressionMatch** suivante active le complément chaque fois qu’un symbole de valeur est inclus dans le corps de l’élément actuel.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## Règle ItemHasKnownEntity


La règle  **ItemHasKnownEntity** active un complément en fonction de l’existence d’une entité dans l’objet ou le corps de l’élément sélectionné. Le type [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) définit les entités prises en charge. L’application d’une expression régulière sur une règle **ItemHasKnownEntity** convient lorsque l’activation est basée sur un sous-ensemble de valeurs pour une entité (par exemple, un ensemble spécifique d’URL, ou des numéros de téléphone avec un certain code régional).


 >
  **Remarque**  Outlook peut extraire des chaînes d’entité en anglais uniquement, indépendamment des paramètres régionaux par défaut spécifiés dans le manifeste. Seuls les messages, mais pas les rendez-vous, peuvent prendre en charge le type d’entité  **MeetingSuggestion**.Vous ne pouvez pas extraire les entités des éléments figurant dans le dossier Éléments envoyés ni utiliser une règle [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) afin d’activer un complément pour les éléments du dossier Éléments envoyés.

La règle  **ItemHasKnownEntity** prend en charge les attributs du tableau suivant. Notez que, bien que la spécification d’une expression régulière soit facultative dans une règle **ItemHasKnownEntity**, si vous choisissez d’utiliser une expression régulière comme filtre d’entité, vous devez spécifier à la fois l’attribut  **RegExFilter** et l’attribut **FilterName**.



|**Nom de l’attribut**|**Description**|
|:-----|:-----|
|**EntityType**|Spécifie le type d’entité qui doit être trouvé pour que la valeur de la règle soit égale à  **true**. Utilisez plusieurs règles pour spécifier plusieurs types d’entité.|
|**RegExFilter**|Spécifie une expression régulière qui filtre les instances de l’entité spécifiée par  **EntityType**.|
|**FilterName**|Spécifie le nom de l’expression régulière spécifiée par  **RegExFilter**, afin qu’il soit possible d’y faire référence ultérieurement par code.|
|**IgnoreCase**|Spécifie s’il faut ignorer la casse pour la correspondance avec l’expression régulière spécifiée par  **RegExFilter**.|

### Exemples

La règle  **ItemHasKnownEntity** suivante active le complément chaque fois qu’une URL se trouve dans l’objet ou le corps de l’élément actuel, et qu’elle contient la chaîne « youtube », indépendamment de la casse de cette chaîne.


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## Utilisation des résultats d’expressions régulières dans le code


Vous pouvez obtenir des correspondances avec une expression régulière en utilisant les méthodes suivantes sur l’élément actif :


- [getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md) renvoie les correspondances de l’élément actuel pour toutes les expressions régulières spécifiées dans les règles **ItemHasRegularExpressionMatch** et **ItemHasKnownEntity** du complément.
    
- [getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md) renvoie les correspondances dans l’élément actuel avec l’expression régulière spécifiée dans une règle **ItemHasRegularExpressionMatch** du complément.
    
- [getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) renvoie les instances complètes des entités qui contiennent des correspondances avec l’expression régulière spécifiée dans une règle **ItemHasKnownEntity** du complément.
    
Lorsque les expressions régulières sont évaluées, les correspondances sont renvoyées vers votre complément dans un objet tableau. Pour  **getRegExMatches**, cet objet a un identifiant correspondant au nom de l’expression régulière. 


 >**Remarque**  Les correspondances renvoyées par un client riche Outlook ne sont pas classées dans un ordre particulier dans le tableau. En outre, vous ne devez pas supposer que le client riche Outlook renvoie les correspondances dans le même ordre que Outlook Web App ou OWA pour périphériques dans ce tableau, même si vous exécutez le même complément sur chacun de ces clients, sur le même élément, et dans la même boîte aux lettres. Pour plus d’informations sur les différences de traitement des expressions régulières entre le client riche Outlook et dans Outlook Web App ou OWA pour périphériques, voir [Limites d’activation et d’API JavaScript des compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).


### Exemples

L’exemple suivant montre un regroupement de règles qui contient une règle  **ItemHasRegularExpressionMatch** avec une expression régulière nommée `videoURL`.


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

L’exemple suivant utilise  **getRegExMatches** dans l’élément actuel pour définir une variable `videos` pour les résultats de la règle **ItemHasRegularExpressionMatch** précédente.




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Plusieurs correspondances sont stockées comme éléments d’un tableau dans cet objet. L’exemple de code suivant montre comment réaliser une itération sur les correspondances pour une expression régulière nommée  `reg1` pour construire une chaîne à afficher sous la forme HTML.




```js
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

L’exemple suivant montre une règle  **ItemHasKnownEntity** qui spécifie l’entité **MeetingSuggestion** et une expression régulière nommée `CampSuggestion`. Outlook active le complément s’il détecte que l’élément sélectionné contient une suggestion de réunion, et que l’objet ou le corps contient le terme « WonderCamp ».




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

L’exemple de code suivant utilise  **getFilteredEntitiesByName(name)** dans l’élément actuel pour définir une variable `suggestions` pour obtenir un tableau des suggestions de réunion détectées pour la règle **ItemHasKnownEntity** précédente.




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## Ressources supplémentaires



- [Créer des compléments Outlook pour des formulaires de lecture](../outlook/read-scenario.md)
    
- [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md)
    
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Meilleures pratiques pour les expressions régulières dans .NET Framework](http://msdn.microsoft.com/en-us/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
