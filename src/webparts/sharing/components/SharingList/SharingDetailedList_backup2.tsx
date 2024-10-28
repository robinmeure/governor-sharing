// import { Text, ShimmeredDetailsList, SearchBox, DefaultButton, Stack, ActionButton, IColumn, SelectionMode, MessageBar, MessageBarType } from '@fluentui/react';
// import * as React from 'react';
// import * as moment from 'moment';
// import { useContext, useEffect, useState, useMemo } from 'react';
// import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
// import { usePnPService } from '../../../../common/services/usePnPService';
// import { Pagination } from '@pnp/spfx-controls-react';
// import { searchQueryGeneratorForDocs } from '../../../../common/utils/Utils';
// import { SearchRequest } from '@microsoft/microsoft-graph-types';
// import { _CONST } from '../../../../common/utils/Const';
// import { useGraphService } from '../../../../common/services/useGraphService';
// import { IDriveItems, IFileSharingResponse, IListItemSearchResponse } from '../../../../common/model';
// import { Toolbar } from '@pnp/spfx-controls-react/lib/Toolbar';
// import { DrivePermissionResponseMapper, GraphSearchResponseMapper } from '../../../../common/config/Mapper';
// import { useBoolean } from '@fluentui/react-hooks';
// import { Person } from '@microsoft/mgt-react';
// import SharedWithColumn from './columnRender/SharedWithColumn';
// import FileExtentionColumn from './columnRender/FileExtentionColumn';
// import LinkColumn from './columnRender/LinkColumn';
// import FileDetailPanel from './panelRender/FileDetailPanel';
// import FilterPanel, { IFilterItem } from './panelRender/FilterPanel';

// interface ISharingDetailedListProps {
//     sharedFiles: IFileSharingResponse[];
// }

// const SharingDetailedList: React.FC<ISharingDetailedListProps> = ({ sharedFiles: initialSharedFiles }): JSX.Element => {
//     const governContext = useContext(SharingWebPartContext);
//     const { getSiteGroups } = usePnPService(governContext.webpartContext);
//     const { getByGraphSearch, getDriveItemsPermission } = useGraphService(governContext.webpartContext);

//     const [sharedFiles, setSharedFiles] = useState<IFileSharingResponse[]>(initialSharedFiles);
//     const [spGroups, setSpGroups] = useState<string[]>();
//     const [selectedDocument, setSelectedDocument] = useState<IFileSharingResponse[]>([]);
//     const [isFilterPanelOpen, { setTrue: openFilterPanel, setFalse: dismissFilterPanel }] = useBoolean(false);
//     const [filterVal, setFilterVal] = useState<IFilterItem>({ siteUrl: "This the test site url", sharedType: [], modifiedBy: "" });
//     const [searchItems, setSearchItems] = useState<IListItemSearchResponse[]>([]);
//     const [currentPage, setCurrentPage] = useState<number>(1);
//     const [totalPages, setTotalPages] = useState<number>(1);
//     const [loading, setLoading] = useState<boolean>(true);
//     const [error, setError] = useState<string>("");

//     const loadPage = async (pageToProcess: number): Promise<void> => {
//         try {
//             setLoading(true);
//             const locFileIds = searchItems.map((item) => item.FileId);
//             const lastIndex = pageToProcess * governContext.pageLimit;
//             const firstIndex = lastIndex - governContext.pageLimit;
//             const paginatedItems = locFileIds.slice(firstIndex, lastIndex);
//             setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));

//             if (paginatedItems.length === 0) {
//                 console.log("No items to display");
//                 return;
//             }

//             const paginatedListItems = paginatedItems.reduce((acc, fileId) => {
//                 const foundItem = searchItems.find(item => item.FileId === fileId);
//                 if (foundItem && foundItem.DriveId) {
//                     acc[fileId] = { driveId: foundItem.DriveId, itemId: foundItem.ItemId };
//                 }
//                 return acc;
//             }, {} as Record<string, IDriveItems>);

//             const locSpGroups = spGroups || await getSiteGroups();
//             if (!spGroups) setSpGroups(locSpGroups);

//             const locSearchItems = searchItems.filter(item => paginatedItems.includes(item.FileId));
//             const driveItems = await getDriveItemsPermission(paginatedListItems);
//             const sharedResults = driveItems.map(driveItem => {
//                 const file = locSearchItems.find(item => item.FileId === driveItem.fileId);
//                 return file ? DrivePermissionResponseMapper(file, driveItem, locSpGroups) : null;
//             }).filter(Boolean) as IFileSharingResponse[];

//             setSharedFiles(sharedResults.filter(result => result.SharedWith !== null));
//         } catch (error) {
//             console.error("Error loading page:", error);
//             setError("Error loading page");
//         }
//         setLoading(false);
//     };

//     const getFilesAndLoadPages = async (queryText: string): Promise<void> => {
//         try {
//             const searchReqForDocs: SearchRequest = {
//                 entityTypes: _CONST.GraphSearch.DocsSearch.EntityType,
//                 query: { queryString: searchQueryGeneratorForDocs(governContext.webpartContext, queryText) },
//                 fields: _CONST.GraphSearch.DocsSearch.Fields,
//                 from: 0,
//                 size: 500
//             };

//             const searchResponse = await getByGraphSearch(searchReqForDocs);
//             const locSearchItems = GraphSearchResponseMapper<IListItemSearchResponse>(searchResponse, _CONST.GraphSearch.DocsSearch.EntityType);
//             setSearchItems(locSearchItems);
//         } catch (error) {
//             console.error("FazLog ~ getFiles ~ error:", error);
//             setLoading(false);
//             setError("Error fetching files");
//         }
//     };

//     useEffect(() => {
//         const init = async (): Promise<void> => {
//             try {
//                 await getFilesAndLoadPages("");
//             } catch (error) {
//                 console.error("FazLog ~ init ~ error:", error);
//             }
//         };
//         init().catch(console.error);
//     }, []);

//     useEffect(() => {
//         if (searchItems.length > 0) {
//             loadPage(1).catch(console.error);
//         }
//     }, [searchItems]);

//     useEffect(() => {
//         if (currentPage) {
//             loadPage(currentPage).catch(console.error);
//         }
//     }, [currentPage]);

//     const columns: IColumn[] = useMemo(() => [
//         {
//             key: 'FileExtension',
//             name: 'FileExtension',
//             fieldName: 'FileExtension',
//             minWidth: 16,
//             maxWidth: 16,
//             isIconOnly: true,
//             isResizable: false,
//             onRender: (item: IFileSharingResponse) => <FileExtentionColumn ext={item.FileExtension} />
//         },
//         {
//             key: 'FileName',
//             name: 'File',
//             fieldName: 'FileName',
//             minWidth: 100,
//             maxWidth: 200,
//             isResizable: true,
//             isSortedDescending: false,
//             isRowHeader: true,
//             sortAscendingAriaLabel: 'Sorted A to Z',
//             sortDescendingAriaLabel: 'Sorted Z to A',
//             data: 'string',
//             onRender: (item: IFileSharingResponse) => <LinkColumn label={item.FileName} url={item.Url} />
//         },
//         {
//             key: 'Channel',
//             name: 'Channel/Folder',
//             fieldName: 'Channel',
//             minWidth: 100,
//             maxWidth: 150,
//             isResizable: true,
//             data: 'string',
//             onRender: (item: IFileSharingResponse) => <LinkColumn label={item.Channel} url={item.FolderUrl} />
//         },
//         {
//             key: 'SharedWith',
//             name: 'Shared',
//             fieldName: 'SharedWith',
//             minWidth: 150,
//             maxWidth: 185,
//             isResizable: true,
//             onRender: (item: IFileSharingResponse) => <SharedWithColumn sharedWith={item.SharedWith} sharedType={item.SharedType} />
//         },
//         {
//             key: 'LastModified',
//             name: 'Modified',
//             fieldName: 'LastModified',
//             minWidth: 70,
//             maxWidth: 90,
//             isResizable: true,
//             isPadded: true,
//             onRender: (item: IFileSharingResponse) => (
//                 <div>
//                     {item.LastModifiedBy?.id && <Person personQuery={item.LastModifiedBy.id} view="oneline" />}
//                     <Text style={{ marginLeft: 36 }} variant="small">{moment(item.LastModified).format('LL')}</Text>
//                 </div>
//             ),
//         },
//         {
//             key: 'SiteUrl',
//             name: 'Site',
//             fieldName: 'SiteUrl',
//             minWidth: 100,
//             maxWidth: 150,
//             isResizable: true,
//             data: 'string',
//             onRender: (item: IFileSharingResponse) => {
//                 const siteName = item.SiteUrl.split("/")[4];
//                 return <LinkColumn label={siteName} url={item.SiteUrl} />;
//             }
//         },
//         {
//             key: "Action",
//             name: "",
//             minWidth: 12,
//             onRender: (item: IFileSharingResponse) => (
//                 <ActionButton iconProps={{ iconName: 'View' }} onClick={() => setSelectedDocument([item])} />
//             )
//         }
//     ], []);

//     if (!loading && error) {
//         return (
//             <div>
//                 <MessageBar messageBarType={MessageBarType.error}>
//                     Something went wrong - {error}
//                 </MessageBar>
//             </div>
//         );
//     }

//     return (
//         <div>
//             <Toolbar actionGroups={{
//                 'share': {
//                     'share': {
//                         title: 'Sharing Settings',
//                         iconName: 'Share',
//                     }
//                 }
//             }}
//                 filters={[{ id: "f1", title: "Guest/External Users" }]}
//                 find={true}
//             />
//             <div>
//                 <Stack horizontal horizontalAlign="space-between">
//                     <Stack.Item grow={3}>
//                         <SearchBox
//                             placeholder="Search..."
//                             underlined={true}
//                             onSearch={async (val) => await getFilesAndLoadPages(val)}
//                             onClear={async () => await getFilesAndLoadPages("")}
//                         />
//                     </Stack.Item>
//                     <Stack horizontalAlign="end" style={{ marginLeft: 12 }}>
//                         <DefaultButton text="Filter" onClick={openFilterPanel} />
//                         <FilterPanel
//                             filterItem={filterVal}
//                             isFilterPanelOpen={isFilterPanelOpen}
//                             onDismissFilterPanel={(newFilter) => {
//                                 if (newFilter) setFilterVal(newFilter);
//                                 dismissFilterPanel();
//                             }}
//                         />
//                     </Stack>
//                 </Stack>
//             </div>

//             {selectedDocument.length > 0 && (
//                 <FileDetailPanel
//                     isOpen={selectedDocument.length > 0}
//                     file={selectedDocument[0]}
//                     onDismiss={() => setSelectedDocument([])}
//                 />
//             )}

//             {!loading && sharedFiles.length === 0 && (
//                 <div>
//                     <MessageBar messageBarType={MessageBarType.info}>
//                         No shared files found.
//                     </MessageBar>
//                 </div>
//             )}

//             <ShimmeredDetailsList
//                 enableShimmer={loading}
//                 usePageCache={true}
//                 columns={columns}
//                 items={sharedFiles}
//                 selectionMode={SelectionMode.none}
//             />

//             <Pagination
//                 key="files"
//                 currentPage={currentPage}
//                 totalPages={totalPages}
//                 onChange={async (page) => setCurrentPage(page)}
//                 limiter={3}
//                 hideFirstPageJump
//                 hideLastPageJump
//                 limiterIcon={"Emoji12"}
//             />
//         </div>
//     );
// };

// export default SharingDetailedList;