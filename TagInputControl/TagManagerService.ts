import { Tag, TagTranslation, Language } from './Tag';
import * as FetchXMLQueries from './FetchXMLQueries';
import { TagManagerServiceError, CustomError } from './ErrorDefinitions';
import * as QueryBuilder from 'odata-query-builder';
import axios, { AxiosResponse } from 'axios';
import Entity = ComponentFramework.WebApi.Entity;
import './ExtensionMethods';


/**
 * The main class of this component. The Tag Manager Service will be responsible to Search Tags, Get a Tag by name, Add and Remove Tags
 * This service uses the webApiProxy from the PCF and also the Web API endpoint from CDS.
 */
export class TagManagerService {

    // just to facilitate the OData WebApi axios request.
    private _webApiDefaultHeaders: any = {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Accept": "application/json"
    };

    // the collection name of the current main entity. Ie: accounts, new_cars, cr983_cats
    private _entitySetName: string;

    // the internal name of the curent entity. Ie: account, new_car, cr983_cat
    private _entityName: string;

    // holds the name of the relationship between the current entity and the tag entity
    private _manyToManyRelationshipName: string;

    // keeps the internal association entity between the Tag and the current Entity.
    private _intersectEntityName:string;

    // the internal unique GUID of the current entity.
    private _entityId: string;

    // the service proxy that will provide access to CDS data.
    private _webApiProxy: ComponentFramework.WebApi;

    // the current user language LCID. Ie.: en-us, es-mx, pt-br.
    private _languageLCID: string;

    // the code of the language that will be used to find the best translation. Ie. 'en', 'es', 'pt'
    private _languageCode: string;

    // the default scope or tag category that will be only used to filter the data.
    private _scope: string;

    // keeps a value of all tags that were already added to prevent duplicate entries.
    private _addedTags: Array<Tag>;

    // the host (URI) of the CRM. Ie: "https://orgb4027a63.crm.dynamics.com"
    private _crmHost: string;

    // keeps the base path to the CRM WebApi
    private _webApiBasePath: string;

    // accessor to the Resources file / strings
    private _resx: ComponentFramework.Resources;

    /**
     * Tag Manager Service default constructor. It will initialize all properties of the Tag Manager Service Instance.
     * @param WebApiProxy The proxy to call all Web API Methods from PCF.
     * @param scope The Tag scope (category).
     * @param languageLCID The LCID of the language of the current user. Ie.: en-us, pt-br.
     * @param searchActivationTreshold Sets the minimum threshold to active the search (DEPRECATED).
     * @param entityName The internal name of the current entity. Ie: account, contact, cr98_training, prefix_entity.
     * @param entityId The internal GUID of the current entity.
     * @param crmHost The default CRM host URI. Ie.: https://orgb4027a63.crm.dynamics.com/.
     * @param manyToManyRelationshipName The custom name of the Many To Many Relationship between the current entity and the Tag entity. If not set, the component will try to load the metadata automatically.
     * @param webApiBasePath The base path to call the Web API of CDS (D365).
     * @param rex The accessor to the strings (resx) file.
     */
    constructor(WebApiProxy: ComponentFramework.WebApi, scope: string, languageLCID: string,
        entityName: string, entityId: string, crmHost: string, manyToManyRelationshipName: string = '', intersectEntityName:string = '', 
        webApiBasePath: string, resx: ComponentFramework.Resources) {

        this._webApiProxy = WebApiProxy;
        this._scope = scope;
        this._languageLCID = languageLCID.toLowerCase();
        this._languageCode = this._languageLCID.split('-').length > 0 ? this._languageLCID.split('-')[0] : 'en';
        this._entityName = entityName;
        this._entityId = entityId;
        this._crmHost = crmHost;
        this._resx = resx;

        if(intersectEntityName != '') {
            this._intersectEntityName = intersectEntityName;
        }

        if (manyToManyRelationshipName != '') {
            this._manyToManyRelationshipName = manyToManyRelationshipName;
        }

        this._webApiBasePath = webApiBasePath;
    }

    /**
     * This method will search for all tags and tags translations based on the given name, the current user language and the defined Tag scope.
     * @param name The filter that will be used on the OData query.
     * @returns {Promise<Array<Tag>>} a list with all Tags and the Translations (as Tags) that were found.
     */
    public async SearchByName(name: string): Promise<Array<Tag>> {

        try {

            if (this._addedTags == null)
                this._addedTags = await this.GetAllByRecordId();

            let query: QueryBuilder.QueryBuilder = new QueryBuilder.QueryBuilder();

            query.select('eax_name, eax_languagelcid, eax_scope')
                .filter(f => f
                    .filterPhrase(`contains(eax_name,'${name}')`)
                    .filterExpression('eax_scope', 'eq', this._scope))
                .orderBy('eax_name');

            let entities: Entity[] = (await this._webApiProxy.retrieveMultipleRecords("eax_tag", query.toQuery())).entities;

            let tags: Array<Tag> = entities.map(e => new Tag(e.eax_name, e.eax_languagelcid, e.eax_scope, this._languageLCID, true, e.eax_tagid));
            let translatedTags: Array<Tag> = await this.SearchTranslationsByName(name);

            // Remove duplicates by name, considering always the current language as the priority.
            let allTags: Array<Tag> = tags.concat(translatedTags).sort(this.SetPreferenceOrder).reduce((acc, item) => {

                // checks if the tag was already added to the current list
                let addedTag: any = this._addedTags.find(i => i.id === item.id);
                if (addedTag == null) {
                    let existingTag: any = acc.find(e => e.name === item.name || e.id === item.id);
                    return (existingTag == null) ? acc.concat(item) : acc;

                } else {
                    return acc;
                }
            }, new Array<Tag>());

            return allTags;

        } catch (err) {
            throw new TagManagerServiceError(this._resx.getString('ErrSearchTagsKey'), err);
        }
    }

    /**
     * This method will only consider those Tags Translations that have the same language of the current user (lcid) or the same LCID code, ie: 'en', 'pt', 'es';
     * the result of this method will be appended to the result of the method SearchByName latter.
     * @param name 
     * @returns {Promise<Array<Tag>>} A list with all Tag Translations (as Tags).
     */
    private async SearchTranslationsByName(name: string): Promise<Array<Tag>> {

        try {

            let query: QueryBuilder.QueryBuilder = new QueryBuilder.QueryBuilder();

            query.select('eax_name, eax_languagelcid, eax_tagscope')
                .expand('eax_Tag($select=eax_tagid)')
                .filter(f => f
                    .filterPhrase(`contains(eax_name,'${name}')`)
                    .filterExpression('eax_tagscope', 'eq', this._scope)
                    .or(f1 => f1
                        .filterExpression('eax_languagelcid', 'eq', this._languageLCID)
                        .filterPhrase(`contains(eax_languagelcid,'${this._languageCode}')`)
                    ), 'and')
                .orderBy('eax_name');

            let entities: Entity[] = (await this._webApiProxy.retrieveMultipleRecords("eax_tagtranslation", query.toQuery())).entities;
            let translatedTags: Array<Tag> = entities.map(e => new Tag(e.eax_name, e.eax_languagelcid, e.eax_scope, this._languageLCID, true, e.eax_Tag.eax_tagid));

            return translatedTags;
        } catch (err) {
            throw new TagManagerServiceError(this._resx.getString('ErrSearchTagTranslationsKey'), err);
        }
    }

    /**
     * This method will try to search for a Tag that has exaclty the same name, language and scope on CDS.
     * This method doesn't consider the Tag Translations.
     * @param name The filter that will be used on the OData query.
     * @returns {Promise<Tag|null>} The Tag record that was found or null in case there is no Tag with the same name, language and scope.
     */
    public async GetByName(name: string): Promise<Tag | null> {
        try {
            let query: QueryBuilder.QueryBuilder = new QueryBuilder.QueryBuilder();
            query.select('eax_name, eax_languagelcid, eax_scope, eax_tagid')
                .top(1)
                .filter(f => f.filterExpression('eax_name', 'eq', name.toLowerCase()));

            let entities: Entity[] = (await this._webApiProxy.retrieveMultipleRecords('eax_tag', query.toQuery())).entities;

            if (entities != null && entities.length > 0) {
                let e: Entity = entities[0];
                return new Tag(e.eax_name, e.eax_languagelcid, e.eax_scope, this._languageLCID, true, e.eax_tagid);
            } else return null;

        } catch (ex) {
            throw new TagManagerServiceError(this._resx.getString('ErrFindTagByNameKey').format({ tagName: name }), ex);
        }
    }

    /**
     * This method will find all Tags (already translated) that are related to the given record.
     * @param recordId the GUID of the record.
     * @returns {Promise<Array<Tag>>} list of all Tags
     */
    public async GetAllByRecordId(): Promise<Array<Tag>> {

        try {

            if(!this._intersectEntityName){
                await this.LoadEntityMetadata();
            }

            let entities: Entity[] = (await this._webApiProxy.retrieveMultipleRecords(this._entityName, "?fetchXml=" + 
            FetchXMLQueries.BuildFindTagsAndTranslationsXML(this._entityName, this._intersectEntityName, this._entityId, this._scope, this._languageLCID))).entities;
           
            this._addedTags = entities.map(t => {

                let tag: Tag = new Tag(t['tags.eax_name'], t['tags.eax_languagelcid'], t['tags.eax_scope'],
                    this._languageLCID, true, t['tags.eax_tagid'], this._entityId);

                // ensuring that the tag will be translated, it keeps the tagid in case the user wants to remove it;
                let translatedName: string = (t['tag_translations.eax_name'] || "");

                if (translatedName != "" && tag.lcid != this._languageLCID) {
                    tag.Translate(translatedName, t['tag_translations.eax_languagelcid'], this._languageLCID);
                }

                return tag;
            });
            
            return this._addedTags;

        } catch (err) {
            throw new TagManagerServiceError(this._resx.getString('ErrGetAllTagsByRecordKey'), err);
        }
    }


    /**
     * Add a tag in case it doesn't exist and append it to the given record. This method will perform the following steps:
     * 1) It will check if the Tag Id is not empty. Easy path: Tag Id != null means that the Tag was selected from the dropdown and the only thing to do is create a relationship between the main entity and the Tag.
	 * 2) If the Tag Id is empty, it will check if any Tag with exactly the same name, scope and language exists on the Tag entity list (CDS).
	 * 3) If the Tag doesn't exist, it will create the Tag on CDS.
     * 4) After the creation, the Tag will be appended to the given record.
     * 
     * @param name The name of the Tag that will be added
     * @param tagId? The internal Tag GUID in case the user selected a Tag from the autocomplete list.
     * @returns {<Promise<Tag | null>} The added tag.
     */
    public async Add(tagName: string, tagId?: string): Promise<Tag | null> {
        try {

            if (this._entitySetName == null) {
                await this.LoadEntityMetadata();
            }

            let tag: Tag | null;

            if (tagId != undefined) {
                tag = new Tag(tagName, this._languageLCID, this._scope, this._languageLCID, false, tagId, this._entityId);
            } else {
                tag = await this.GetByName(name);
                if (!tag || (tag && !tag.id)) {
                    tag = await this.Create(tagName);
                }
            }

            if (tag && tag.id) {
                let relateResult: boolean = await this.Relate(this._entityId, tag.id, this._entitySetName, this._manyToManyRelationshipName, 'eax_tags');
                if (relateResult) {
                    tag.persisted = true;
                    this._addedTags.push(tag);
                } else throw new CustomError(this._resx.getString('ErrTagRelationshipKey').format({ tagName: tagName }));
            } else throw new CustomError(this._resx.getString('ErrTagCreationKey').format({ tagName: tagName }));

            return tag;

        } catch (err) {
            throw new TagManagerServiceError(this._resx.getString('ErrTagCreationRelationshipKey'), err);
        }
    }

    /**
     * This is an internal and helper method that correlates any primary entity with the secondary entity. This method can be used to correlate entities on a 1:many or many:many relationship
     * Since there is no WebApi method on the PCF proxy to create relationships, this method uses the standard Web API Common Data Service REST endpoint.  
     * @see https://docs.microsoft.com/en-us/previous-versions/dynamicscrm-2016/developers-guide/mt607875(v=crm.8)?redirectedfrom=MSDN#bkmk_Removeareferencetoanentity
     * @param primaryEntityId The internal GUID of the main entity.
     * @param relatedEntityId The internal GUID of the related entity.
     * @param primaryEntityName The internal name of the main entity. Ie: account, new_car, cr983_category.
     * @param relationshipName The internal name of the relationship between these entities. It is a case sensitive name.
     * @param relatedEntityName The internal name of the related entity. Ie: account, new_car, cr983_category.
     * @returns {Promise<boolean>} true if success | false if not
     */
    private async Relate(primaryEntityId: string, relatedEntityId: string, primaryEntitySetName: string, relationshipName: string, relatedEntitySetName: string): Promise<boolean> {

        let response: AxiosResponse<any> = await axios.post(`${this._webApiBasePath}/${primaryEntitySetName}(${primaryEntityId})/${relationshipName}/$ref`,
            {
                '@odata.id': `${this._crmHost}${this._webApiBasePath}/${relatedEntitySetName}(${relatedEntityId})`
            },
            { headers: this._webApiDefaultHeaders });

        return (response.status == 204);
    }


    /**
     * This is a internal and helper method that will create a Tag and related the created Tag with the current user language.
     * In case the language doesn't exist on CDS, and exception will be thrown.
     * @param tagName The name of the Tag that will be created.
     * @returns {Promise<Tag | null>} The created Tag instance.
     */
    private async Create(tagName: string): Promise<Tag | null> {

        let tag: Tag | null = null;

        // it will create a new Tag and returns the last added GUID
        let addedTagId: any = (await this._webApiProxy.createRecord('eax_tag', {
            'eax_name': tagName,
            'eax_scope': this._scope
        })).id;

        if (!addedTagId) {
            throw new CustomError(this._resx.getString('ErrTagCreationKey').format({ tagName: tagName }));
        }

        let language: Language | null = await this.GetLanguageByLCID(this._languageLCID);

        if (!language || (language && !language.id)) {
            throw new CustomError(this._resx.getString('ErrTagCreatedLCIDNotFoundKey').format({ tagName: tagName, languageLCID: this._languageLCID }));
        }

        const setLanguageResponse: boolean = await this.Relate(addedTagId, language.id, 'eax_tags', 'eax_Language', 'eax_languages');

        if (!setLanguageResponse) {
            throw new CustomError(this._resx.getString('ErrTagCreatedLanguageNotSetKey').format({ tagName: tagName, languageLCID: this._languageLCID }));
        } else {
            tag = new Tag(tagName, this._languageLCID, this._scope, this._languageLCID, false, addedTagId, this._entityId);
        }

        return tag;
    }


    /**
     * This method will ONLY remove the association between the given entity and the Tag (TagId). It will not delete the Tag from the Tags list.
     * @param tagId The internal GUID of the Tag that will be removed.
     * @returns {<Promise<boolean>} true if the Tag was removed | false if not.
     */
    public async Remove(tagId: string): Promise<boolean> {
        try {

            if (this._entityId != null && tagId != null) {
                if (this._entitySetName == null) {
                    await this.LoadEntityMetadata();
                }

                let deleteResponse: AxiosResponse<any> = await axios.delete(`${this._webApiBasePath}/${this._entitySetName}(${this._entityId})/${this._manyToManyRelationshipName}/$ref?$id=${this._webApiBasePath}/eax_tags(${tagId})`,
                    {
                        headers: this._webApiDefaultHeaders
                    });

                this._addedTags = this._addedTags.filter(i => i.id != tagId);

                return (deleteResponse.status == 204); // tag association removed, 204 status = no content.

            } else return false;
        } catch (err) {
            throw new TagManagerServiceError(this._resx.getString('ErrTagRemovalKey'), err);
        }
    }

    /**
     * This method will check if the given Tag has the same language that the current user language. If so, it will update the main Tag record
     * If not, it will update or create a Tag Translation record.
     * @param tagId The internal GUID of the Tag that will be updated
     * @param name The new name of the Tag
     * @returns {Promise<Tag>} The updated Tag instace
     */
    public async Update(tagId: string, name: string): Promise<Tag> {
        throw new Error('Method Not Implemented Error.');
    }


    /**
     * This is a internal and helper method that will try to load the main EntitySetName (ie.:accounts, contacts) and the name of the many to many relationship
     * between the main entity and the Tag entity. 
     */
    private async LoadEntityMetadata() {

        this._entitySetName = (await axios.get(`${this._webApiBasePath}/EntityDefinitions(LogicalName='${this._entityName}')?$select=EntitySetName`)).data.EntitySetName;

        if (this._entitySetName == null) {
            throw new CustomError(this._resx.getString('ErrEntitySetNotFoundKey').format({ entityName: this._entityName }));
        }

        // just in case it is a custom many to many relationship. Applicable when there are more than one M:M relationship with the Tag entity.
        if (this._manyToManyRelationshipName == null) {
            let relationshipNames: Array<any> = (await axios.get(`${this._webApiBasePath}/EntityDefinitions(LogicalName='${this._entityName}')/ManyToManyRelationships?$select=SchemaName,IntersectEntityName&$filter=Entity1LogicalName eq 'eax_tag'`)).data.value;

            if (!relationshipNames || (relationshipNames && (!relationshipNames.length || (relationshipNames.length > 0 && (!relationshipNames[0].SchemaName || !relationshipNames[0].IntersectEntityName )) ))) {
                throw new CustomError(this._resx.getString('ErrManyToManyRelationshipNotFoundKey').format({ entityName: this._entityName }));
            }

            this._manyToManyRelationshipName = relationshipNames[0].SchemaName;
            this._intersectEntityName = relationshipNames[0].IntersectEntityName;
        }
    }

    /**
     * Internal an helper method. It gets the first record of a language entity based on the LCID.
     * @param lcid The language LCID standard (ie.: en-us, pt-br, es-mx)
     * @returns {Promise<Language | null>} The language entity record.
     */
    private async GetLanguageByLCID(lcid: string): Promise<Language | null> {
        try {
            let query: QueryBuilder.QueryBuilder = new QueryBuilder.QueryBuilder();

            query.select('eax_name, eax_lcid, eax_code, eax_languageid')
                .top(1)
                .filter(f => f.filterExpression('eax_lcid', 'eq', lcid))

            let entities: Entity[] = (await this._webApiProxy.retrieveMultipleRecords('eax_language', query.toQuery())).entities;

            if (entities != null && entities.length > 0) {
                let e: Entity = entities[0];
                return new Language(e.eax_languageid, e.eax_lcid, e.eax_code, e.eax_name);

            } return null;

        } catch (ex) {
            throw new TagManagerServiceError(this._resx.getString('ErrFindLCIDKey').format({ lcid : lcid }), ex);
        }
    }

    /**
     * Tag preferece comparator to decide which tag comes first based on the current user language;
     * Based on the Tag Language and User Language, there is a definition of preferences. @see Tag class
     * @param tag1 The first Tag to be compared
     * @param tag2 The Tag it will compared to
     * @returns {number} 0 if equals or error, 1 if Tag1.preference is higher than Tag2, -1 if Tag2 is higher than Tag1.preference
     */
    private SetPreferenceOrder(tag1: Tag, tag2: Tag): number {

        if (tag1 == null || tag2 == null)
            return 0;

        if (tag1.preference > tag2.preference)
            return 1;

        if (tag1.preference < tag2.preference)
            return -1;

        return 0;

    }
}