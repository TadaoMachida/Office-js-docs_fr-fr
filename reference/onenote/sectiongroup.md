# Objet SectionGroup (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_   


Représente un groupe de sections OneNote. Les groupes de sections peuvent contenir des sections et des groupes de sections.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|clientUrl{|chaîne|URL du client du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-clientUrl{)|
|id|string|Obtient l’ID du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-id)|
|name|chaîne|Obtient le nom du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-name)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|notebook|[Bloc-notes](notebook.md)|Obtient le bloc-notes qui contient le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-notebook)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Obtient le groupe de sections qui contient le groupe de sections. Génère ItemNotFound si le groupe de sections est un enfant direct du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|Obtient le groupe de sections qui contient le groupe de sections. Renvoie la valeur Null si le groupe de sections est un enfant direct du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroupOrNull)|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|Collection de groupes de sections dans le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sectionGroups)|
|Sections|[SectionCollection](sectioncollection.md)|Collection de sections dans le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sections)|

## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[addSection(title: String)](#addsectiontitle-string)|[Section](section.md)|Ajoute une nouvelle section à la fin du groupe de sections.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Ajoute un nouveau groupe de sections à la fin de cet objet sectionGroup.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-load)|

## Détails des méthodes


### addSection(title: String)
Ajoute une nouvelle section à la fin du groupe de sections.

#### Syntaxe
```js
sectionGroupObject.addSection(title);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|title|String|Nom de la nouvelle section.|

#### Retourne
[Section](section.md)

#### Exemples
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;
    
    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Add a section to each section group.
            $.each(sectionGroups.items, function(index, sectionGroup) {
                sectionGroup.addSection("Agenda");
            });
            
            // Run the queued commands.
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### addSectionGroup(name: String)
Ajoute un nouveau groupe de sections à la fin de cet objet sectionGroup.

#### Syntaxe
```js
sectionGroupObject.addSectionGroup(name);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|String|Nom de la nouvelle section.|

#### Retourne
[SectionGroup](sectiongroup.md)

#### Exemples
```js          
OneNote.run(function (context) {
    var sectionGroup;
    var nestedSectionGroup;

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroups = notebook.sectionGroups;

    // Queue a command to load the new section group.
    sectionGroups.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function(){
            sectionGroup = sectionGroups.items[0];
            sectionGroup.load();
            return context.sync();
        })
        .then(function(){
            nestedSectionGroup = sectionGroup.addSectionGroup("Sample nested section group");
            nestedSectionGroup.load();
            return context.sync();
        })
        .then(function() {
            console.log("New nested section group name is " + nestedSectionGroup.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Retourne
void
### Exemples d’accès aux propriétés

**id**
```js
OneNote.run(function (context) {
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group. 
    // For best performance, request specific properties.           
    sectionGroup.load("id,name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Section group ID: " + sectionGroup.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name and notebook**
```js
OneNote.run(function (context) {
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group with the specified properties.           
    sectionGroup.load("name,notebook/name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Parent notebook name: " + sectionGroup.notebook.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sectionGroups**
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("name");
    
    // Get the child section groups of the first section group in the notebook.
    var nestedSectionGroups = sectionGroups._GetItem(0).sectionGroups;
    
    // Queue a command to load the ID and name properties of the child section groups.
    nestedSectionGroups.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each child section group.
            $.each(nestedSectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
            });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**Sections**
```js
OneNote.run(function (context) {

    // Get the sections that are siblings of the current section.
    var sections = context.application.getActiveSection().parentSectionGroup.sections;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sections.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each section.
            $.each(sections.items, function(index, section) {
                console.log("Section name: " + section.name);  
                console.log("Section ID: " + section.id);  
            });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

