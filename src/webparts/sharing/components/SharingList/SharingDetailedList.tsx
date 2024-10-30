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
            sharedType: ["Link", "Inherited", "Member", "Guest", "Everyone", "Group"],
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

            // const locSpGroups: string[] = await pnpService.getSiteGroups();

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

    // const localDataFilter = (updatedFilter: IPaginationFilterState, filterType: ("SharedType" | "User")[]): void => {
    //     try {
    //         const isSharedTypeFilter = filterType.filter(val => val === "SharedType").length > 0;
    //         if (isSharedTypeFilter) {
    //             // const locSharedFiles = sharedFiles.filter(val => updatedFilter.filterVal.sharedType.indexOf(val.SharedType) > -1);
    //             const locSharedFiles = sharedFiles.filter(val => val.SharedWith.filter(val2 => updatedFilter.filterVal.sharedType.indexOf(val2.type) > -1).length > 0);
    //             setSharedFiles(locSharedFiles);
    //         }
    //     } catch (error) {
    //     }
    // }

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
            onRender: (item: IFileSharingResponse) => <SharedWithColumn
                sharedWith={item.SharedWith}
                sharedType={item.SharedType}
                filteredSharedTypes={paginationFilterState.filterVal.sharedType} />
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
    ], [paginationFilterState.filterVal.sharedType]);

    if (!loadingErrorState.loading && loadingErrorState.error) {
        return (
            <div>
                <MessageBar messageBarType={MessageBarType.error}>
                    Something went wrong - {loadingErrorState.error}
                </MessageBar>
            </div>
        );
    }
    // const handleSearchTermDebounce = (inputValue: string): void => {
    //     setPaginationFilterState(prevState => ({
    //         ...prevState,
    //         searchQuery: inputValue,
    //         isRefreshData: true
    //     }));
    // }
    // const debounceSearchTerm = useCallback(_.debounce(handleSearchTermDebounce, 2000), []);
    // const handleChange = (newValue: string): void => {
    //     debounceSearchTerm(newValue);
    // };

    return (
        <div>
            <div>
                <Stack horizontal horizontalAlign="space-between">
                    <Stack.Item grow={3}>

                        <div style={{ maxWidth: "800px" }}>

                            <SearchBox
                                placeholder="Search..."
                                underlined={true}
                                // onChange={async (_e, val) => {
                                //     if (val) {
                                //         handleChange(val);
                                //     } else {
                                //         setPaginationFilterState(prevState => ({
                                //             ...prevState,
                                //             searchQuery: "",
                                //             isRefreshData: true
                                //         }));
                                //     }
                                // }}
                                onSearch={async (val: string) => {
                                    setPaginationFilterState(prevState => ({
                                        ...prevState,
                                        searchQuery: val,
                                        isRefreshData: true
                                    }));
                                }}
                                onClear={async () => {
                                    setPaginationFilterState(prevState => ({
                                        ...prevState,
                                        searchQuery: "",
                                        isRefreshData: true
                                    }));
                                }}
                            />
                        </div>

                    </Stack.Item>
                    <Stack horizontalAlign="end" style={{ marginLeft: 12 }}>
                        <PrimaryButton iconProps={{ iconName: 'Filter' }}
                            text="Filter" onClick={openFilterPanel} />

                        {isFilterPanelOpen &&
                            <FilterPanel
                                filterItem={paginationFilterState.filterVal}
                                isFilterPanelOpen={isFilterPanelOpen}
                                onDismissFilterPanel={async (newFilter) => {
                                    dismissFilterPanel();
                                    if (newFilter) {
                                        //check for server refresh
                                        const isServerRefresh = newFilter.siteUrl !== paginationFilterState.filterVal.siteUrl;
                                        const updatedFilter: IPaginationFilterState = {
                                            ...paginationFilterState,
                                            filterVal: newFilter,
                                            currentPage: 1,
                                            isRefreshData: isServerRefresh
                                        };
                                        // if (!isServerRefresh) {
                                        //     localDataFilter(updatedFilter, ["SharedType"]);
                                        // }
                                        setPaginationFilterState(updatedFilter);
                                    }
                                }}
                            />
                        }

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