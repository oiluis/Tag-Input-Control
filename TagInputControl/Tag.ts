/**
 * Holds all tag information
 */
export class Tag{

    /**
     * The internal GUID of the tag
     */
    public readonly id?:string;

    /**
     * Equivalent to Tag.name
     */
    public value:string;

    /**
     * The name of the Tag. Ie: Power Apps, Power Automate, Sales etc.
     */
    public name :string;

    /**
     * Any Tag Category, such as skills, colors, aspects etc.
     */
    public readonly scope: string;

    /**
     * The language LCID of this Tag, ie.: en-us, pt-br, es-mx. 
     */
    public lcid : string;

    /**
     * The related record internal GUID. This will hold the main entity GUID.
     */
    public recordId? : string;

    /**
     * A list with all Tag Translations (not used)
     */
    public translations : Array<TagTranslation>;

    /**
     * Defines the preference to display the Tag on the autocomplete. The preference index is calculated based on the user language
     */
    public preference : number;

    /**
     * Defines if a Tag is persisted or not on CDS
     */
    public persisted : boolean;

    /**
     * Keeps the original LCID after a translation.
     */
    public originallcid : string;


    /**
     * Default constructor, only sets the properties and calculates the Preference Index.
     * @param name The name of the Tag. Ie: Power Apps, Power Automate, Sales etc.
     * @param lcid The language LCID of this Tag, ie.: en-us, pt-br, es-mx. 
     * @param scope Any Tag Category, such as skills, colors, aspects etc.
     * @param userlcid The current user languaged LCID, ie.: en-us, pt-br, es-mx. 
     * @param persisted Defines if a Tag is persisted or not on CDS
     * @param tagId The internal GUID of the tag
     * @param recordId The related record internal GUID. This will hold the main entity GUID.
     */
    constructor(name:string, lcid:string, scope:string, userlcid:string, persisted:boolean, tagId?:string,  recordId? : string){
        this.value = name;
        this.name = name;
        this.scope = scope;
        this.lcid = lcid.toLowerCase();
        this.recordId = recordId;
        this.id = tagId;
        this.persisted = persisted;
        this.CalculatePreference(userlcid);
    }

    /**
     * This method will only change the name and the current Tag and the LCID. It will also keeps the original LCID for the update action.
     * @param name The translated name of the tag.
     * @param lcid The language that was used to display / translate the name.
     * @param userlcid The user language used to translated the Tag.
     */
    public Translate(name: string, lcid:string, userlcid:string){
        this.name = name;
        this.value = name;
        this.originallcid = this.lcid;
        this.lcid = lcid.toLowerCase();
        this.CalculatePreference(userlcid);
    }

    /**
     * Sets the preference index base on the user language. This will be used later to display the Tags to the user and remove duplicated entries.
     * @param userlcid The user language used to translated the Tag.
     */
    private CalculatePreference(userlcid:string) : void{

        if(this.lcid == userlcid){
            this.preference = 1;
        } else  if(this.lcid.split('-')[0] === userlcid.split('-')[0]){
            this.preference = 2;
        } else {
            this.preference = 3;
        }
    }
    
}

/**
 * (not used) It will be used by the update process ... Still not implemented
 */
export class TagTranslation{
    
    public id?:string;
    public tagId:string;
    public value:string;
    public name:string;
    public lcid : string;

    constructor(tagId:string, lcid:string, name: string){
        this.tagId = tagId;
        this.lcid = lcid;
        this.name = name;
        this.value = name;
    }
}

export class Language {
    public id:string;
    public lcid:string;
    public code:string;
    public name:string;

    constructor(id:string,lcid:string,code:string,name:string){
        this.id = id;
        this.lcid = lcid;
        this.code = code;
        this.name = name;
    }
}