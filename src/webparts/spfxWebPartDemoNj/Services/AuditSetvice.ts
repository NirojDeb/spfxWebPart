import * as React from 'react';
import { Comments, Item, sp } from "@pnp/sp/presets/all";
import "@pnp/sp/sites";
import { IContextInfo } from "@pnp/sp/sites";
import { Web } from "@pnp/sp/webs";
import axios from 'axios';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import "@pnp/sp/lists";

export class AuditService{
    public getListItems()
    {
        return sp.web.lists.getByTitle('AuditManagerTest').items.getAll();
    }

    public getAttachmentsbyId(id:number)
    {
        return sp.web.lists.getByTitle('AuditManagerTest').items.getById(id).select('Attachments').expand('AttachmentFiles').get()
    }

    public createAudit(audit)
    {
        return sp.web.lists.getByTitle('AuditManagerTest').items.add(audit);
    }

    public addAttachementsToAudit(auditId,files)
    {
        return sp.web.lists.getByTitle('AuditManagerTest').items.getById(auditId).attachmentFiles.addMultiple(files);
    }
}