import { Text, ShimmeredDetailsList, SearchBox, DefaultButton, Stack, ActionButton, IColumn, SelectionMode, MessageBar, MessageBarType } from '@fluentui/react';
import * as React from 'react'
import * as moment from 'moment';
import { useContext, useEffect, useState } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { usePnPService } from '../../../../common/services/usePnPService';
import { Pagination } from '@pnp/spfx-controls-react';
import { searchQueryGeneratorForDocs } from '../../../../common/utils/Utils';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { _CONST } from '../../../../common/utils/Const';
import { useGraphService } from '../../../../common/services/useGraphService';
import { IDriveItems, IFileSharingResponse, IListItemSearchResponse } from '../../../../common/model';
import { Toolbar } from '@pnp/spfx-controls-react/lib/Toolbar';
import { DrivePermissionResponseMapper, GraphSearchResponseMapper } from '../../../../common/config/Mapper';
import { useBoolean } from '@fluentui/react-hooks';
import { Person } from '@microsoft/mgt-react';
import SharedWithColumn from './columnRender/SharedWithColumn';
import FileExtentionColumn from './columnRender/FileExtentionColumn';
import LinkColumn from './columnRender/LinkColumn';
import FileDetailPanel from './panelRender/FileDetailPanel';
import FilterPanel, { IFilterItem } from './panelRender/FilterPanel';

const SharingDetailedList: React.FC = (): JSX.Element => {

    const governContext = useContext(SharingWebPartContext);

    const { getSiteGroups
    } = usePnPService(governContext.webpartContext);
    const { getByGraphSearch, getDriveItemsPermission } = useGraphService(governContext.webpartContext);

    const [sharedFiles, setSharedFiles] = useState<IFileSharingResponse[]>([]);
    const [fileIds, setFileIds] = useState<string[]>([]);
    const [spGroups, setSpGroups] = useState<string[]>();

    const [selectedDocument, setSelectedDocument] = useState<IFileSharingResponse[]>([]);

    //const [searchKeyword, setSearchKeyword] = useState<string>("");
    const [isFilterPanelOpen, { setTrue: openFilterPanel, setFalse: dismissFilterPanel }] = useBoolean(false);
    const [filtreVal, setFilterVal] = useState<IFilterItem>({
        siteUrl: "This the test site url",
        sharedType: [],
        modifiedBy: ""
    });

    // let searchItems: Record<string, any> = [];
    const [searchItems, setSearchItems] = useState<IListItemSearchResponse[]>([]);
    const [currentPage, setCurrentPage] = useState<number>();
    const [totalPages, setTotalPages] = useState<number>(1);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string>("");

    const loadPage = async (paramFileIds: string[], pageToProcess: number): Promise<void> => {
        try {
            setCurrentPage(pageToProcess);
            const lastIndex = pageToProcess ? pageToProcess * governContext.pageLimit : 1;
            const firstIndex = lastIndex - governContext.pageLimit;
            const paginatedItems = paramFileIds.slice(firstIndex, lastIndex);
            setTotalPages(Math.ceil(paramFileIds.length / governContext.pageLimit));

            if (paginatedItems.length === 0) {
                console.log("No items to display");
                return;
            }

            const paginatedListItems = paginatedItems.reduce((acc, fileId) => {
                const foundItem = searchItems.filter(item => item.FileId === fileId);
                if (foundItem.length === 0 || !foundItem[0].DriveId || !foundItem[0].DriveId) return null;
                acc[fileId] = {
                    driveId: foundItem[0].DriveId || '',
                    itemId: foundItem[0].ItemId || ''
                };
                return acc;
            }, {} as Record<string, IDriveItems>);

            //TODO handle this differently
            const locSpGroups: string[] = spGroups ? spGroups : await getSiteGroups();
            if (spGroups === undefined) {
                setSpGroups(locSpGroups);
            }
            // const sharedLinkResults = await getSharingLinks(searchItems, locSpGroups);

            // get searchItems where fileIds are in paginatedItems
            const locSearchItems = searchItems.filter(item => paginatedItems.includes(item.FileId));
            const sharedResults: IFileSharingResponse[] = [];
            // const driveItemParam = locSearchItems.map(item => ({ driveId: item.DriveId, driveItemId: item.DriveItemId }));
            if (paginatedListItems) {
                const driveItems = await getDriveItemsPermission(paginatedListItems);

                // now we have all the data we need, we can start building up the result
                driveItems.forEach(driveItem => {
                    const file = locSearchItems.filter(item => item.FileId === driveItem.fileId)[0];
                    const locSharedResult = DrivePermissionResponseMapper(file, driveItem, locSpGroups);
                    sharedResults.push(locSharedResult);
                });

                if (!sharedResults) {
                    return;
                }
                const sharingLinks = sharedResults.filter(result => result.SharedWith !== null);
                setSharedFiles(sharingLinks);
            } else {
                throw new Error("Paginated list items are null");
            }

        } catch (error) {
            console.error("Error loading page:", error);
            setError("Error loading page");
        }
    };

    const getFilesAndLoadPages = async (queryText: string): Promise<void> => {
        try {
            setLoading(true);


            const searchReqForDocs: SearchRequest = {
                entityTypes: _CONST.GraphSearch.DocsSearch.EntityType,
                query: {
                    queryString: `${searchQueryGeneratorForDocs(governContext.webpartContext, queryText)}`
                },
                fields: _CONST.GraphSearch.DocsSearch.Fields,
                from: 0,
                size: 500
            };

            const searchResponse = await getByGraphSearch(searchReqForDocs);
            const locSearchItems = GraphSearchResponseMapper<IListItemSearchResponse>(searchResponse, _CONST.GraphSearch.DocsSearch.EntityType);
            console.log("FazLog ~ init:", locSearchItems);



            setSearchItems(locSearchItems);
            // get all file ids
            const locFileIds = locSearchItems.map((item) => item.FileId);
            await loadPage(locFileIds, 1);
            setFileIds(locFileIds);

            setLoading(false);

            // await loadPage(locFileIds);

            // return locFileIds;
        } catch (error) {
            console.log("FazLog ~ getFiles ~ error:", error);
            setLoading(false);
            setError(error);
            throw error;
        }
    }

    // const _requestSharingDetails = async (): Promise<void> => {
    //     try {
    //         setLoading(true);
    //         await getFilesAndLoadPages("");
    //         // await loadPage(locFileIds, 1);
    //         setLoading(false);
    //     } catch (error) {
    //         console.log("FazLog ~ error:", error);
    //         throw error;
    //     }
    // }

    // useEffect(() => {
    //     if (currentPage !== undefined) {
    //         loadPage(fileIds).catch(error => console.error("Error loading page:", error));
    //     }
    // }, [currentPage]);

    useEffect(() => {
        const init = async (): Promise<void> => {

            try {
                await getFilesAndLoadPages("");
                // setLoading(true);
                // const locFileIds = await getFiles("");
                // setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));
                // setCurrentPage(1);
                // await loadPage(locFileIds);
            } catch (error) {
                console.log("FazLog ~ init ~ error:", error);
            }
        };
        init().catch(error => console.error("Error during initialization:", error));
    }, []);

    const _columns: IColumn[] = [
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
            minWidth: 70,
            maxWidth: 90,
            isResizable: true,
            isPadded: true,
            onRender: (item: IFileSharingResponse) => {
                return <div>
                    {/* {item.LastModifiedBy?.displayName} */}
                    {item.LastModifiedBy && item.LastModifiedBy.id && <Person personQuery={item.LastModifiedBy.id} view="oneline" />}
                    <Text style={{ marginLeft: 36 }} variant="small">{moment(item.LastModified).format('LL')}</Text>
                    {/* <br />
                    {format(new Date(item.LastModified), 'dd-MMM-yyyy')} */}
                </div>
            },
        },
        // {
        //     key: 'SharingUserType',
        //     name: 'SharingUserType',
        //     fieldName: 'SharingUserType',
        //     minWidth: 16,
        //     maxWidth: 16,
        //     isIconOnly: true,
        //     isResizable: false
        // },
        {
            key: 'SiteUrl',
            name: 'Site',
            fieldName: 'SiteUrl',
            minWidth: 100,
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
            minWidth: 12,
            onRender: (item: IFileSharingResponse) => {
                return <ActionButton iconProps={{ iconName: 'View' }} onClick={() => setSelectedDocument([item])} />
            }
        }
    ];

    if (!loading && error) {
        return <div>
            <MessageBar messageBarType={MessageBarType.error}>
                Something went wrong - {error}
            </MessageBar>
        </div>
    }

    return (
        <div>
            <Toolbar actionGroups={{
                'share': {
                    'share': {
                        title: 'Sharing Settings',
                        iconName: 'Share',
                    }
                }
            }}
                filters={[
                    {
                        id: "f1",
                        title: "Guest/External Users",
                    }]}
                find={true}
            />
            <div>
                <Stack horizontal horizontalAlign="space-between">
                    <Stack.Item grow={3}>
                        <SearchBox
                            placeholder="Search..."
                            underlined={true}
                            //value={searchKeyword}
                            // onChange={(e, newVal) => setSearchKeyword(newVal ? newVal : "")}
                            onSearch={async (val) => {
                                console.log("FazLog ~ val:", val);
                                // const locFileIds = await getFiles(val);
                                // setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));
                                // setCurrentPage(1);
                                // await loadPage(locFileIds);

                                await getFilesAndLoadPages(val);
                            }}
                            onClear={async () => {

                                // const locFileIds = await getFiles("");
                                // setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));
                                // setCurrentPage(1);
                                // await loadPage(locFileIds);

                                await getFilesAndLoadPages("");
                            }}
                        />
                    </Stack.Item>
                    <Stack horizontalAlign="end" style={{ marginLeft: 12 }}>
                        <DefaultButton text="Filter" onClick={openFilterPanel} />
                        <FilterPanel
                            filterItem={filtreVal}
                            isFilterPanelOpen={isFilterPanelOpen}
                            onDismissFilterPanel={(newFilter) => {
                                if (newFilter) {
                                    setFilterVal(newFilter);
                                }
                                dismissFilterPanel();
                            }} />
                    </Stack>
                </Stack>
            </div>

            {selectedDocument?.length > 0 &&
                <FileDetailPanel
                    isOpen={selectedDocument?.length > 0}
                    file={selectedDocument[0]}
                    onDismiss={() => {
                        // dismissFileDetailPanel();
                        setSelectedDocument([]);
                    }
                    } />
            }

            {!loading && sharedFiles.length === 0 && <div>
                <MessageBar messageBarType={MessageBarType.info}>
                    { }
                </MessageBar>
            </div>}


            <ShimmeredDetailsList
                enableShimmer={loading}
                usePageCache={true}
                columns={_columns}
                items={sharedFiles}
                selectionMode={SelectionMode.none}
            />


            <Pagination
                key="files"
                currentPage={currentPage || 1}
                totalPages={totalPages}
                onChange={async (page) => {
                    await loadPage(fileIds, page);
                    // setCurrentPage(page);
                }}
                limiter={3}
                hideFirstPageJump
                hideLastPageJump
                limiterIcon={"Emoji12"}
            />
        </div>
    )
}

export default SharingDetailedList;