import { Web } from "@pnp/sp/presets/all";
import { sp } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ConsoleListener, Logger, LogLevel } from "@pnp/logging";
import { ISonaEmdProps } from "../../components/ISonaEmdProps";
// import { SPHttpClient } from 'http-client';

export interface ISPCRUDOPS {
    getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, props: ISonaEmdProps): Promise<any>;
    getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, props: ISonaEmdProps): Promise<any>;
    insertData(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    updateData(listName: string, itemId: number, data: any, props: ISonaEmdProps): Promise<any>;
    deleteData(listName: string, itemId: number, props: ISonaEmdProps): Promise<any>;
    getListInfo(listName: string, props: ISonaEmdProps): Promise<any>;
    getListData(listName: string, columnsToRetrieve: string, props: ISonaEmdProps): Promise<any>;
    batchInsert(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    batchUpdate(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    batchDelete(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    createFolder(listName: string, folderName: string, props: ISonaEmdProps): Promise<any>;
    uploadFile(folderServerRelativeUrl: string, file: File, props: ISonaEmdProps): Promise<any>;
    deleteFile(fileServerRelativeUrl: string, props: ISonaEmdProps): Promise<any>;
    currentProfile(props: ISonaEmdProps): Promise<any>;
    //currentUserProfile(props: IDeviationuatProps): Promise<any>;
    getLoggedInSiteGroups(props: ISonaEmdProps): Promise<any>;
    getAllSiteGroups(props: ISonaEmdProps): Promise<any>;
    getTopData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
        , orderby: { column: string, isAscending: boolean }, top: number, props: ISonaEmdProps): Promise<any>;
    addAttchmentInList(data: File, listName: string, itemId: number, fileName: string, props: ISonaEmdProps): Promise<any>;
}

class SPCRUDOPSImpl implements ISPCRUDOPS {
    async getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, props: ISonaEmdProps): Promise<any> {
        if (!props.currentSPContext || !props.currentSPContext.pageContext) {
            throw new Error('SharePoint context is not available');
        }
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        let items = web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) {
            items = items.select(columnsToRetrieve);
        }
        if (columnsToExpand) {
            items = items.expand(columnsToExpand);
        }
        if (filters) {
            items = items.filter(filters);
        }
        if (orderby) {
            items = items.orderBy(orderby.column, orderby.isAscending);
        }
        return await items.getAll();
    }

    async getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, props: ISonaEmdProps): Promise<any> {
        if (!props.currentSPContext || !props.currentSPContext.pageContext) {
            throw new Error('SharePoint context is not available');
        }
        const fullUrl = props.currentSPContext.pageContext.web.absoluteUrl;
        const parts = fullUrl.split('/');
        const baseUrl = parts.slice(0, 5).join('/');
        const web = Web(baseUrl);
        let items = web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) {
            items = items.select(columnsToRetrieve);
        }
        if (columnsToExpand) {
            items = items.expand(columnsToExpand);
        }
        if (filters) {
            items = items.filter(filters);
        }
        if (orderby) {
            items = items.orderBy(orderby.column, orderby.isAscending);
        }
        return await items.getAll();
    }
    async insertData(listName: string, data: any, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).items.add(data);
    }

    async updateData(listName: string, itemId: number, data: any, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).items.getById(itemId).update(data);
    }

    async deleteData(listName: string, itemId: number, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).items.getById(itemId).recycle();
    }

    async getListInfo(listName: string, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).get();
    }

    async getListData(listName: string, columnsToRetrieve: string, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        let items = web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) {
            items = items.select(columnsToRetrieve);
        }
        return await items.get();
    }

    async batchInsert(listName: string, data: any, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        const entityTypeFullName = await web.lists.getByTitle(listName).getListItemEntityTypeFullName();
        const batch = web.createBatch();
        data.forEach((item: any) => {
            web.lists.getByTitle(listName).items.inBatch(batch).add(item, entityTypeFullName);
        });
        return await batch.execute();
    }

    async batchUpdate(listName: string, data: any, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        const batch = web.createBatch();
        data.forEach((item: any) => {
            web.lists.getByTitle(listName).items.getById(item.Id).inBatch(batch).update(item);
        });
        return await batch.execute();
    }

    async batchDelete(listName: string, data: any, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        const batch = web.createBatch();
        data.forEach((item: any) => {
            web.lists.getByTitle(listName).items.getById(item.Id).inBatch(batch).delete();
        });
        return await batch.execute();
    }

    async createFolder(listName: string, folderName: string, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).rootFolder.folders.addUsingPath(folderName);
    }

    async uploadFile(folderServerRelativeUrl: string, file: File, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.name, file, true);
    }

    async deleteFile(fileServerRelativeUrl: string, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.getFileByServerRelativeUrl(fileServerRelativeUrl).recycle();
    }

    // async currentProfile(props: ISonaEmdProps): Promise<any> {
    //     const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
    //     return await web.currentUser.get();
    // }

    async currentProfile(props: ISonaEmdProps): Promise<any> {
        try {
            const profile = await sp.profiles.myProperties.get();
            return profile;
        } catch (error) {
            console.error("Error fetching user profile", error);
            return null;
        }
    }

    async getLoggedInSiteGroups(props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.currentUser.groups.get();
    }

    async getAllSiteGroups(props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.siteGroups.get();
    }

    async getTopData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
        , orderby: { column: string, isAscending: boolean }, top: number, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        let items = web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) {
            items = items.select(columnsToRetrieve);
        }
        if (columnsToExpand) {
            items = items.expand(columnsToExpand);
        }
        if (filters) {
            items = items.filter(filters);
        }
        if (orderby) {
            items = items.orderBy(orderby.column, orderby.isAscending);
        }
        if (top) {
            items = items.top(top);
        }
        return await items.get();
    }

    async addAttchmentInList(data: File, listName: string, itemId: number, fileName: string, props: ISonaEmdProps): Promise<any> {
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        return await web.lists.getByTitle(listName).items.getById(itemId).attachmentFiles.add(fileName, data);
    }
}

export default function SPCRUDOPS(): Promise<ISPCRUDOPS> {
    return Promise.resolve(new SPCRUDOPSImpl());
}