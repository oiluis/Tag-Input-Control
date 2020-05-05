export const BuildFindTagsAndTranslationsXML = (enityName:string, intersectEntityName:string, entityId:string, scope:string, lcdi:string) =>
 `<fetch>
    <entity name='${enityName}'>
      <filter type='and'>
        <condition attribute='${enityName}id' operator='eq' value='${entityId}' />
      </filter>
      <link-entity name='${intersectEntityName}' from='${enityName}id' to='${enityName}id' intersect='true'>
        <link-entity name='eax_tag' from='eax_tagid' to='eax_tagid' alias='tags'>
          <attribute name='eax_name' />
          <attribute name='eax_languagelcid' />
          <attribute name='eax_tagid' />
          <attribute name='eax_scope' />
          <filter type='and'>
            <condition attribute='eax_scope' operator='eq' value='${scope}' />
          </filter>
          <link-entity name='eax_tagtranslation' from='eax_tag' to='eax_tagid' link-type='outer' alias='tag_translations'>
            <attribute name='eax_name' />
            <attribute name='eax_tag' />
            <attribute name='eax_languagelcid' />
            <attribute name='eax_tagscope' />
            <filter type='and'>
              <condition attribute='eax_tagscope' operator='eq' value='${scope}' />
              <condition attribute='eax_languagelcid' operator='eq' value='${lcdi}' />
            </filter>
          </link-entity>
        </link-entity>
      </link-entity>
    </entity>
  </fetch>`.replace(/(\r\n|\n|\r)/gm, "");