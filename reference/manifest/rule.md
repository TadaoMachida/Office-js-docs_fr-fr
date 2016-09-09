
# Élément Rule
Spécifie les règles d’activation à évaluer pour ce complément de messagerie.

 **Type de complément :** messagerie


## Syntaxe :

 **ItemIs Rule** - Définit une règle qui donne la valeur True si l’élément sélectionné est du type spécifié.


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **ItemHasAttachment Rule** - Définit une règle qui donne la valeur True si l’élément contient une pièce jointe.




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity** - Définit une règle qui donne la valeur True si l’élément contient dans son objet ou son corps du texte correspondant au type d’entité spécifié.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **ItemHasRegularExpressionMatch Rule** - Définit une règle qui donne la valeur True si une correspondance de l’expression régulière spécifiée est trouvée dans la propriété spécifiée de l’élément.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlaintext" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **RuleCollection Rule** - Définit une collection de règles et l’opérateur logique à utiliser lors de leur évaluation.




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## Contenu dans :

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Attributs :

 **Attributs ItemIs Rule**



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType (chaîne)|obligatoire|Spécifie le type d’élément à mettre en correspondance. Les options disponibles sont :

|**ItemType**|**Élément ItemClass correspondant**|
|:-----|:-----|
|Rendez-vous|IPM.Appointment|
|Message(1)|Inclut les messages électroniques, les demandes, les réponses et les annulations de réunion.|
|
|FormType|ItemFormType (chaîne)|obligatoire|Spécifie si l’application doit apparaître dans le formulaire de lecture ou de modification pour l’élément. Les options disponibles peuvent être l’une des suivantes :|

|**FormType**|**Description**|
|:-----|:-----|
|Lecture|Indique qu’il faut activer le complément de messagerie uniquement dans les formulaires de lecture (de l’élément **ItemType** indiqué).|
|Modifier|Indique qu’il faut activer le complément de messagerie uniquement dans les formulaires de composition (de l’élément **ItemType** indiqué).|
|ReadOrEdit|Indique qu’il faut activer le complément de messagerie dans les formulaires de lecture et de composition (de l’élément **ItemType** indiqué).|
|ItemClass|chaîne|facultatif|Spécifie la classe de message personnalisé à mettre en correspondance. Pour plus d’informations, voir l’article relatif à l’[activation d’un complément de messagerie dans Outlook pour une classe de message spécifique](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx).|
|IncludeSubClasses|booléen|facultatif|Spécifie si la règle doit donner la valeur true si l’élément est une sous-classe de la classe de message spécifiée ; par défaut, la valeur est false.|


(1) Les éléments suivants sont les classes de message correspondantes : IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled.

 **Attributs de la règle ItemHasAttachment**

Aucun.

 **Attributs ItemHasKnownEntity Rule**



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType (chaîne)|obligatoire|Spécifie le type d’entité à rechercher pour que la règle donne la valeur True. Il peut s’agir de l’un des éléments suivants :

|**KnownEntityType**|**Description**|
|:-----|:-----|
|MeetingSuggestion|Texte identifié par reconnaissance de modèle comme étant une référence à un événement ou une réunion.|
|TaskSuggestion| Texte identifié par reconnaissance de modèle comme contenant une expression pouvant donner lieu à une action.|
|Address|Texte identifié par reconnaissance de modèle comme étant une référence à une adresse postale aux États-Unis.|
|Url|Texte identifié par reconnaissance de modèle comme contenant un nom de fichier ou une URL d’adresse web.|
|PhoneNumber| Série de chiffres identifiée par reconnaissance de modèle comme étant un numéro de téléphone en Amérique du Nord.|
|EmailAddress|Texte identifié par reconnaissance de modèle comme contenant une adresse de messagerie au format SMTP.|
|Contact|Texte identifié par reconnaissance de modèle comme contenant des informations de contact.|
|RegExFilter|chaîne|facultatif|Spécifie une expression régulière à exécuter par rapport à cette entité à des fins d’activation.|
|FilterName|chaîne|facultatif|Spécifie le nom du filtre d’expression régulière, afin qu’il soit possible par la suite de s’y référer dans le code de votre complément.|
|IgnoreCase|booléen|facultatif|Indique d’ignorer la casse lors de l’exécution de l’expression régulière spécifiée par l’attribut **RegExFilter**.|
 **Attributs ItemHasRegularExpressionMatch Rule**



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|RegExName|chaîne|obligatoire|Spécifie le nom de l’expression régulière afin que vous puissiez vous référer à l’expression dans le code de votre complément.|
|RegExValue|chaîne|obligatoire|Spécifie l’expression régulière qui sera évaluée pour déterminer si le complément de messagerie doit être affiché. |
|PropertyName|PropertyName (chaîne)|obligatoire|Spécifie le nom de la propriété par rapport à laquelle l’expression sera évaluée. Les options disponibles sont :

|**PropertyName**|**Description**|
|:-----|:-----|
|Objet|Évalue l’expression régulière par rapport à l’objet de l’élément.|
|BodyAsPlaintext|Évalue l’expression régulière par rapport au corps de l’élément en texte brut.|
|BodyAsHtml|Évalue l’expression régulière par rapport au corps de l’élément si le corps est disponible en HTML.|
|SenderSTMPAddress|Évalue l’expression régulière par rapport à l’adresse SMTP de l’expéditeur de l’élément.|
|IgnoreCase|booléen|facultatif|Indique d’ignorer la casse lors de l’exécution de l’expression régulière.|
 **Attributs RuleCollection Rule**



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|Mode|string|obligatoire|Spécifie l’opérateur logique à utiliser lors de l’évaluation de cette collection de règles. Il peut s’agir des éléments suivants : « And » ou « Or ».|

## Ressources supplémentaires



- 
  [Activer un complément de messagerie dans Outlook pour une classe de message spécifique](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) et [Règles d’activation pour les compléments Outlook](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)
    
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
