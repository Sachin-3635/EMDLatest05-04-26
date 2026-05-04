import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { IPatelEngProps } from "../../components/IPatelEngProps";
import { ISonaEmdProps } from "../../components/ISonaEmdProps";
import SPCRUDOPS from "../../service/DAL/spcrudops";
 
export interface ISPCRUD {
    [x: string]: any;
    getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonaEmdProps): Promise<any>;
    getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonaEmdProps): Promise<any>;
    insertData(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    updateData(listName: string, itemId: number, data: any, props: ISonaEmdProps): Promise<any>;
    deleteData(listName: string, itemId: number, props: ISonaEmdProps): Promise<any>;
    getListInfo(listName: string, props: ISonaEmdProps): Promise<any>;
    getListData(listName: string, columnsToRetrieve: string, props: ISonaEmdProps): Promise<any>;
    batchInsert(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    batchUpdate(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    batchDelete(listName: string, data: any, props: ISonaEmdProps): Promise<any>;
    createFolder(listName: string, folderName: string, props: ISonaEmdProps):Promise<any>;
    uploadFile(folderServerRelativeUrl: string, file: File, props: ISonaEmdProps): Promise<any>;
    deleteFile(fileServerRelativeUrl: string, props: ISonaEmdProps): Promise<any>;
    currentProfile(props: ISonaEmdProps): Promise<any>;
    //currentUserProfile(props: IDeviationuatProps): Promise<any>;
    getLoggedInSiteGroups(props: ISonaEmdProps): Promise<any>;
    getAllSiteGroups(props: ISonaEmdProps): Promise<any>;
    getTopData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, top: number, props: ISonaEmdProps): Promise<any>;
    addAttchmentInList(attFiles: File, listName: string, itemId: number, fileName: string, props: ISonaEmdProps): Promise<any>;
}

export default async function USESPCRUD(): Promise<ISPCRUD> {
    const spCrudOps = await SPCRUDOPS();
    return {
        getData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonaEmdProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        getRootData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonaEmdProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        insertData: async (listName: string, data: any, props: ISonaEmdProps) => {
            return await spCrudOps.insertData(listName, data, props);
        },
        updateData: async (listName: string, itemId: number, data: any, props: ISonaEmdProps) => {
            return await spCrudOps.updateData(listName, itemId, data, props);
        },
        deleteData: async (listName: string, itemId: number, props: ISonaEmdProps) => {
            return await spCrudOps.deleteData(listName, itemId, props);
        },
        getListInfo: async (listName: string, props: ISonaEmdProps) => {
            return await spCrudOps.getListInfo(listName, props);
        },
        getListData: async (listName: string, columnsToRetrieve: string, props: ISonaEmdProps) => {
            return await spCrudOps.getListData(listName, columnsToRetrieve, props);
        },
        batchInsert: async (listName: string, data: any, props: ISonaEmdProps) => {
            return await spCrudOps.batchInsert(listName, data, props);
        },
        batchUpdate: async (listName: string, data: any, props: ISonaEmdProps) => {
            return await spCrudOps.batchUpdate(listName, data, props);
        },
        batchDelete: async (listName: string, data: any, props: ISonaEmdProps) => {
            return await spCrudOps.batchDelete(listName, data, props);
        },
        createFolder: async (listName: string, folderName: string, props: ISonaEmdProps) => {
            return await spCrudOps.createFolder(listName, folderName, props);
        },
        uploadFile: async (folderServerRelativeUrl: string, file: File, props: ISonaEmdProps) => {
            return await spCrudOps.uploadFile(folderServerRelativeUrl, file, props);
        },
        deleteFile: async (fileServerRelativeUrl: string, props: ISonaEmdProps) => {
            return await spCrudOps.deleteFile(fileServerRelativeUrl, props);
        },
        currentProfile: async (props: ISonaEmdProps) => {
            return await spCrudOps.currentProfile(props);
        },
        // const currentUserProfile = async (props: IDeviationuatProps) => {
          
        //    // const queryUrl = "https://etgworld.sharepoint.com/sites/UAT_BPM/_api/web/currentuser/groups";
            
        //     const result: any = await (await spCrudOps).currentUserProfile( props);
        //     return result;
        // };
        getLoggedInSiteGroups: async (props: ISonaEmdProps) => {
            return await spCrudOps.getLoggedInSiteGroups(props);
        },
        getAllSiteGroups: async (props: ISonaEmdProps) => {
            return await spCrudOps.getAllSiteGroups(props);
        },
        getTopData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean }, top: number, props: ISonaEmdProps) => {
            return await spCrudOps.getTopData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, top, props);
        },
        addAttchmentInList: async (attFiles: File, listName: string, itemId: number, fileName: string, props: ISonaEmdProps) => {
            return await spCrudOps.addAttchmentInList(attFiles, listName, itemId, fileName, props);
        }
    };
}