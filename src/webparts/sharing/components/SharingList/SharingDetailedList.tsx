import { ShimmeredDetailsList, SelectionMode, MessageBar, MessageBarType } from '@fluentui/react';
import * as React from 'react';
import { useContext, useEffect, useState } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { Pagination } from '@pnp/spfx-controls-react';
import { searchQueryGeneratorForFiles } from '../../../../common/utils/Utils';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { _CONST } from '../../../../common/utils/Const';
import { useGraphService } from '../../../../common/services/useGraphService';
import { IDriveItems, IFileSharingResponse, IGraphResponseMetadata, IListItemSearchResponse } from '../../../../common/model';
import { DrivePermissionResponseMapper, GraphSearchResponseMapper, GraphSearchResultMetadata } from '../../../../common/utils/Mapper';
import FilePermissionPanel from './itemDetail/FilePermissionPanel';
import { IFilterItem } from './filter/FilterPanel';
import FilterItems from './filter/FilterItems';
import Columns from './columnRender/Columns';


interface ILoadingErrorState {
    loading: boolean;
    error: string;
}

export interface IPaginationFilterState {
    searchKeyword: string;
    queryString: string;
    currentPage: number;
    totalPages: number;
    searchMetadata: IGraphResponseMetadata | undefined;
    filterVal: IFilterItem;
    selectedDocument: IFileSharingResponse | undefined;
    isRefreshData: boolean
}

const SharingDetailedList: React.FC = (): JSX.Element => {
    const governContext = useContext(SharingWebPartContext);
    const webpartContext = governContext.webpartContext;
    const { getByGraphSearch, getDriveItemsPermission } = useGraphService(webpartContext);

    const [loadingErrorState, setLoadingErrorState] = useState<ILoadingErrorState>({
        loading: true,
        error: ""
    });

    const [paginationFilterState, setPaginationFilterState] = useState<IPaginationFilterState>({
        currentPage: 1,
        totalPages: 1,
        filterVal: {
            siteUrl: governContext.isTeams ? webpartContext.pageContext.site.absoluteUrl : "",
            sharedType: ["Link", "Inherited", "Member", "Guest", "Everyone", "Group"],
            modifiedBy: "",
            fileFolder: "BothFilesFolders",
            preQuery: governContext.webpartProperties.preQuery
        }
        ,
        searchMetadata: undefined,
        selectedDocument: undefined,
        searchKeyword: "",
        queryString: "",
        isRefreshData: false
    });

    const [searchItems, setSearchItems] = useState<IListItemSearchResponse[]>([]);
    const [sharedFiles, setSharedFiles] = useState<IFileSharingResponse[]>([]);

    const getFiles = async (queryFilter: IPaginationFilterState): Promise<IListItemSearchResponse[]> => {
        try {
            const queryString = searchQueryGeneratorForFiles(webpartContext, queryFilter);
            const searchReqForDocs: SearchRequest & { trimDuplicates: boolean } = {
                entityTypes: _CONST.GraphSearch.DocsSearch.EntityType,
                query: { queryString: queryString },
                fields: _CONST.GraphSearch.DocsSearch.Fields,
                from: 0,
                size: 500,
                trimDuplicates: true
            };

            const searchResponse = await getByGraphSearch(searchReqForDocs);
            const searchMetadata = GraphSearchResultMetadata(searchResponse);
            if (searchMetadata) {
                setPaginationFilterState(prevState => ({
                    ...prevState,
                    searchMetadata: searchMetadata,
                    queryString: queryString
                }));
            }
            const locSearchItems = GraphSearchResponseMapper<IListItemSearchResponse>(searchResponse, _CONST.GraphSearch.DocsSearch.EntityType);
            return locSearchItems;
        } catch (error) {
            setLoadingErrorState(prevState => ({ ...prevState, error: "Error fetching files", loading: false }));
            throw error;
        }
    };

    const loadPage = async (pageToProcess: number, searchItems: IListItemSearchResponse[]): Promise<void> => {
        try {
            if (!loadingErrorState.loading) {
                setLoadingErrorState(prevState => ({ ...prevState, loading: true }));
            }
            const locFileIds = searchItems.map((item) => item.FileId);
            const lastIndex = pageToProcess * governContext.pageLimit;
            const firstIndex = lastIndex - governContext.pageLimit;
            const paginatedItems = locFileIds.slice(firstIndex, lastIndex);
            setPaginationFilterState(prevState => ({ ...prevState, totalPages: Math.ceil(locFileIds.length / governContext.pageLimit) }));

            // const locSpGroups: string[] = await getSiteGroups(paginationFilterState.filterVal.siteUrl);

            if (paginatedItems.length === 0) {
                setSharedFiles([]);
                setLoadingErrorState(prevState => ({ ...prevState, loading: false }));
                return;
            }
            const paginatedListItems = paginatedItems.reduce((acc, fileId) => {
                const foundItem = searchItems.find((item: IListItemSearchResponse) => item.FileId === fileId);
                if (foundItem && foundItem.DriveId) {
                    acc[fileId] = { driveId: foundItem.DriveId, itemId: foundItem.ItemId };
                }
                return acc;
            }, {} as Record<string, IDriveItems>);

            const locSearchItems = searchItems.filter(item => paginatedItems.includes(item.FileId));
            const driveItemsPermissions = await getDriveItemsPermission(paginatedListItems);
            const sharedResults = driveItemsPermissions.map(driveItem => {
                const file = locSearchItems.find(item => item.FileId === driveItem.fileId);
                if (file) {
                    const locDrivePermission = DrivePermissionResponseMapper(file, driveItem);
                    if (paginationFilterState.filterVal.sharedType.length > 0) {
                        locDrivePermission.SharedWith = locDrivePermission.SharedWith.filter((val) => paginationFilterState.filterVal.sharedType.includes(val.type))
                    }
                    return locDrivePermission;
                }
                return null;
            }).filter(Boolean) as IFileSharingResponse[];

            // check if paginationFilterState.filterVal.sharedType is not empty
            // if (paginationFilterState.filterVal.sharedType.length > 0) {
            //     const filteredSharedResults = sharedResults.filter(result => {
            //         if (result.SharedWith.length > 0) {
            //             return result.SharedWith.some(val => paginationFilterState.filterVal.sharedType.includes(val.type));
            //         }
            //         return false;
            //     });
            //     setSharedFiles(filteredSharedResults);
            // } else {
            //     setSharedFiles(sharedResults.filter(result => result.SharedWith !== null));
            // }
            setSharedFiles(sharedResults.filter(result => result.SharedWith !== null));
            setLoadingErrorState(prevState => ({ ...prevState, loading: false }));
        } catch (error) {
            console.error("Error loading page:", error);
            setLoadingErrorState(prevState => ({ ...prevState, error: "Error loading page", loading: false }));
        }
    };

    const getFilesAndLoadPages = async (preFilterPageVal: IPaginationFilterState): Promise<void> => {
        try {
            setLoadingErrorState(prevState => ({ ...prevState, loading: true }));
            const locSearchVals = await getFiles(preFilterPageVal);
            setSearchItems(locSearchVals);
            await loadPage(1, locSearchVals);
            setLoadingErrorState(prevState => ({ ...prevState, loading: false }));
        } catch (error) {
            setLoadingErrorState(prevState => ({ ...prevState, error: "Error fetching files", loading: false }));
            throw error;
        }
    };

    useEffect(() => {
        const init = async (): Promise<void> => {
            await getFilesAndLoadPages(paginationFilterState);
        };
        init().catch(error => console.error("Error during initialization:", error));
    }, []);

    useEffect(() => {
        const serverRefreshData = async (): Promise<void> => {
            if (paginationFilterState.isRefreshData) {
                await getFilesAndLoadPages(paginationFilterState);
                setPaginationFilterState(prevState => ({
                    ...prevState,
                    isRefreshData: false
                }));
            }
        };
        serverRefreshData().catch(error => console.error("Error during refreshData:", error));
    }, [paginationFilterState.isRefreshData])

    if (!loadingErrorState.loading && loadingErrorState.error) {
        return (
            <div>
                <MessageBar messageBarType={MessageBarType.error}>
                    Something went wrong - {loadingErrorState.error}
                </MessageBar>
            </div>
        );
    }

    return (
        <div>
            <FilterItems paginationFilterState={paginationFilterState} setPaginationFilterState={setPaginationFilterState} />

            {paginationFilterState.selectedDocument && (
                <FilePermissionPanel
                    isOpen={!!paginationFilterState.selectedDocument}
                    file={paginationFilterState.selectedDocument}
                    onDismiss={() => setPaginationFilterState(prevState => ({ ...prevState, selectedDocument: undefined, isRefreshData: false }))}
                />
            )}

            {!loadingErrorState.loading && sharedFiles.length === 0 && (
                <div style={{ padding: "12px 0" }}>
                    <MessageBar messageBarType={MessageBarType.info}>
                        No shared files found.
                    </MessageBar>
                </div>
            )}

            {(loadingErrorState.loading || sharedFiles.length > 0) && <>
                <ShimmeredDetailsList
                    enableShimmer={loadingErrorState.loading}
                    usePageCache={true}
                    columns={Columns({
                        paginationFilterState,
                        setPaginationFilterState,
                    })}
                    items={sharedFiles}
                    selectionMode={SelectionMode.none}
                />

                <Pagination
                    key="files"
                    currentPage={paginationFilterState.currentPage || 1}
                    totalPages={paginationFilterState.totalPages}
                    onChange={async (page) => {
                        setPaginationFilterState(prevState => ({ ...prevState, currentPage: page, isRefreshData: false }));
                        await loadPage(page, searchItems)
                    }
                    }
                    limiter={4}
                    hideFirstPageJump
                    hideLastPageJump
                />
            </>}


        </div>
    );
};

export default SharingDetailedList;