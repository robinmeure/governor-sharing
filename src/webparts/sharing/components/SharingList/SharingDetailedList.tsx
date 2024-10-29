import { Text, ShimmeredDetailsList, SearchBox, Stack, ActionButton, IColumn, SelectionMode, MessageBar, MessageBarType, Persona, PersonaSize, PrimaryButton } from '@fluentui/react';
import * as React from 'react';
import * as moment from 'moment';
import { useContext, useEffect, useState, useMemo } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { Pagination } from '@pnp/spfx-controls-react';
import { searchQueryGeneratorForDocs } from '../../../../common/utils/Utils';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { _CONST } from '../../../../common/utils/Const';
import { useGraphService } from '../../../../common/services/useGraphService';
import { IDriveItems, IFileSharingResponse, IListItemSearchResponse } from '../../../../common/model';
import { DrivePermissionResponseMapper, GraphSearchResponseMapper } from '../../../../common/config/Mapper';
import { useBoolean } from '@fluentui/react-hooks';
import SharedWithColumn from './columnRender/SharedWithColumn';
import FileExtentionColumn from './columnRender/FileExtentionColumn';
import LinkColumn from './columnRender/LinkColumn';
import FileDetailPanel from './panelRender/FileDetailPanel';
import FilterPanel, { IFilterItem } from './panelRender/FilterPanel';


interface ILoadingErrorState {
    loading: boolean;
    error: string;
}

export interface IPaginationFilterState {
    searchQuery: string;
    currentPage: number;
    totalPages: number;
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
            siteUrl: webpartContext.pageContext.site.absoluteUrl,
            sharedType: [],
            modifiedBy: ""
        },
        selectedDocument: undefined,
        searchQuery: "",
        isRefreshData: false
    });


    const [searchItems, setSearchItems] = useState<IListItemSearchResponse[]>([]);
    const [sharedFiles, setSharedFiles] = useState<IFileSharingResponse[]>([]);

    const [isFilterPanelOpen, { setTrue: openFilterPanel, setFalse: dismissFilterPanel }] = useBoolean(false);

    const getFiles = async (query: IPaginationFilterState): Promise<IListItemSearchResponse[]> => {
        try {
            const searchReqForDocs: SearchRequest | {} = {
                entityTypes: _CONST.GraphSearch.DocsSearch.EntityType,
                query: { queryString: searchQueryGeneratorForDocs(webpartContext, query) },
                fields: _CONST.GraphSearch.DocsSearch.Fields,
                from: 0,
                size: 500,
                trimDuplicates: true
            };

            const searchResponse = await getByGraphSearch(searchReqForDocs);
            console.log("FazLog ~ getFiles ~ searchResponse:", searchResponse);
            const locSearchItems = GraphSearchResponseMapper<IListItemSearchResponse>(searchResponse, _CONST.GraphSearch.DocsSearch.EntityType);
            return locSearchItems;
        } catch (error) {
            console.error("FazLog ~ getFiles ~ error:", error);
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
            setPaginationFilterState(prevState => ({ ...prevState, totalPages: Math.ceil(locFileIds.length / governContext.pageLimit), isRefreshData: false }));

            // const locSpGroups: string[] = await getSiteGroups(paginationFilterState.filterVal.siteUrl);
            // console.log("FazLog ~ loadPage ~ locSpGroups:", locSpGroups);

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

            console.log("FazLog ~ searchItems ~ searchItems:", searchItems);
            console.log("FazLog ~ paginatedListItems ~ paginatedListItems:", paginatedListItems);

            // const locSpGroups: string[] = await pnpService.getSiteGroups();
            // console.log("FazLog ~ loadPage ~ locSpGroups:", locSpGroups);

            const locSearchItems = searchItems.filter(item => paginatedItems.includes(item.FileId));
            const driveItemsPermissions = await getDriveItemsPermission(paginatedListItems);
            console.log("FazLog ~ loadPage ~ driveItemsPermissions:", driveItemsPermissions);
            const sharedResults = driveItemsPermissions.map(driveItem => {
                const file = locSearchItems.find(item => item.FileId === driveItem.fileId);
                if (file) {
                    const locDrivePermission = DrivePermissionResponseMapper(file, driveItem);
                    console.log("FazLog ~ sharedResults ~ locDrivePermission:", locDrivePermission);
                    if (paginationFilterState.filterVal.sharedType.length > 0) {
                        locDrivePermission.SharedWith = locDrivePermission.SharedWith.filter((val) => paginationFilterState.filterVal.sharedType.includes(val.type))
                    }
                    return locDrivePermission;
                }
                return null;
            }).filter(Boolean) as IFileSharingResponse[];

            setSharedFiles(sharedResults.filter(result => result.SharedWith !== null));
            setLoadingErrorState(prevState => ({ ...prevState, loading: false }));
        } catch (error) {
            console.error("Error loading page:", error);
            setLoadingErrorState(prevState => ({ ...prevState, error: "Error loading page", loading: false }));
        }
    };

    const getFilesAndLoadPages = async (query: IPaginationFilterState): Promise<void> => {
        try {
            setLoadingErrorState(prevState => ({ ...prevState, loading: true }));
            const locSearchVals = await getFiles(query);
            setSearchItems(locSearchVals);
            await loadPage(1, locSearchVals);
            setLoadingErrorState(prevState => ({ ...prevState, loading: false }));
        } catch (error) {
            console.error("FazLog ~ getFilesAndLoadPages ~ error:", error);
            setLoadingErrorState(prevState => ({ ...prevState, error: "Error fetching files", loading: false }));
        }
    };



    useEffect(() => {
        const init = async (): Promise<void> => {
            try {
                await getFilesAndLoadPages(paginationFilterState);
            } catch (error) {
                console.error("FazLog ~ init ~ error:", error);
            }
        };
        init().catch(error => console.error("Error during initialization:", error));
    }, []);

    useEffect(() => {
        const refreshData = async (): Promise<void> => {
            try {
                if (paginationFilterState.isRefreshData) {
                    await getFilesAndLoadPages(paginationFilterState);
                    setPaginationFilterState(prevState => ({
                        ...prevState,
                        isRefreshData: false
                    }));
                } else {
                    // do offline filter 
                }
            } catch (error) {
                console.error("FazLog ~ refreshData ~ error:", error);
            }
        };
        refreshData().catch(error => console.error("Error during refreshData:", error));
    }, [paginationFilterState.isRefreshData])


    const columns: IColumn[] = useMemo(() => [
        {
            key: 'FileExtension',
            name: 'FileExtension',
            fieldName: 'FileExtension',
            minWidth: 16,
            maxWidth: 16,
            isIconOnly: true,
            isResizable: false,
            onRender: (item: IFileSharingResponse) => <FileExtentionColumn ext={item.FileExtension} />
        },
        {
            key: 'FileName',
            name: 'File',
            fieldName: 'FileName',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
            isSortedDescending: false,
            isRowHeader: true,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            onRender: (item: IFileSharingResponse) => <LinkColumn label={item.FileName} url={item.Url} />
        },
        {
            key: 'Channel',
            name: 'Channel/Folder',
            fieldName: 'Channel',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            data: 'string',
            onRender: (item: IFileSharingResponse) => <LinkColumn label={item.Channel} url={item.FolderUrl} />
        },
        {
            key: 'SharedWith',
            name: 'Shared',
            fieldName: 'SharedWith',
            minWidth: 150,
            maxWidth: 185,
            isResizable: true,
            onRender: (item: IFileSharingResponse) => <SharedWithColumn sharedWith={item.SharedWith} sharedType={item.SharedType} />
        },

        {
            key: 'LastModified',
            name: 'Modified',
            fieldName: 'LastModified',
            minWidth: 120,
            maxWidth: 170,
            isResizable: true,
            isPadded: true,
            onRender: (item: IFileSharingResponse) => (
                <div>
                    {item.LastModifiedBy?.id &&
                        <Persona
                            size={PersonaSize.size24}
                            imageAlt={item.LastModifiedBy?.displayName || ''}
                            imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${item.LastModifiedBy?.id}`}
                            text={item.LastModifiedBy?.displayName || ''}
                            secondaryText={item.LastModifiedBy?.id}
                        />
                    }
                    <Text style={{ marginLeft: 32 }} variant="small">{moment(item.LastModified).format('LL')}</Text>
                </div>
            ),
        },
        {
            key: 'SiteUrl',
            name: 'Site',
            fieldName: 'SiteUrl',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            data: 'string',
            onRender: (item: IFileSharingResponse) => {
                const siteName = item.SiteUrl.split("/")[4];
                return <LinkColumn label={siteName} url={item.SiteUrl} />;
            }
        },
        {
            key: "Action",
            name: "",
            minWidth: 4,
            onRender: (item: IFileSharingResponse) => (
                <ActionButton iconProps={{ iconName: 'View' }} onClick={() => setPaginationFilterState(prevState => ({ ...prevState, selectedDocument: item }))} />
            )
        }
    ], []);

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
            <div>
                <Stack horizontal horizontalAlign="space-between">
                    <Stack.Item grow={3}>
                        <SearchBox
                            placeholder="Search..."
                            underlined={true}
                            onSearch={async (val: string) => {
                                // const updatedFilter: IPaginationFilterState = {
                                //     ...paginationFilterState,
                                //     searchQuery: val,
                                //     isRefreshData: true
                                // };
                                setPaginationFilterState(prevState => ({
                                    ...prevState,
                                    searchQuery: val,
                                    isRefreshData: true
                                }));
                                // setPaginationFilterState(updatedFilter);
                                // await getFilesAndLoadPages(updatedFilter);
                            }}
                            onClear={async () => {
                                // const updatedFilter: IPaginationFilterState = {
                                //     ...paginationFilterState,
                                //     searchQuery: "",
                                //     isRefreshData: true
                                // };
                                // setPaginationFilterState(updatedFilter);

                                setPaginationFilterState(prevState => ({
                                    ...prevState,
                                    searchQuery: "",
                                    isRefreshData: true
                                }));
                                // await getFilesAndLoadPages(updatedFilter);
                            }}
                        />
                    </Stack.Item>
                    <Stack horizontalAlign="end" style={{ marginLeft: 12 }}>
                        <PrimaryButton text="Filter" onClick={openFilterPanel} />
                        <FilterPanel
                            filterItem={paginationFilterState.filterVal}
                            isFilterPanelOpen={isFilterPanelOpen}
                            onDismissFilterPanel={async (newFilter) => {
                                dismissFilterPanel();
                                if (newFilter) {
                                    console.log("FazLog ~ onDismissFilterPanel={ ~ newFilter:", newFilter);
                                    const updatedFilter: IPaginationFilterState = {
                                        ...paginationFilterState,
                                        filterVal: newFilter,
                                        currentPage: 1,
                                        isRefreshData: true
                                    };
                                    await getFilesAndLoadPages(updatedFilter);
                                    // setPaginationFilterState(updatedFilter)
                                }
                            }}
                        />
                    </Stack>
                </Stack>
            </div>

            {paginationFilterState.selectedDocument && (
                <FileDetailPanel
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
                    columns={columns}
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