/* eslint-disable */

import { Facepile, IColumn, Icon, Link, OverflowButtonType, Persona, Text, PersonaSize, ShimmeredDetailsList, TooltipHost, SelectionMode, DialogType, MarqueeSelection } from '@fluentui/react';
import * as React from 'react'
import * as moment from 'moment';
import { FileIconType, getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { useContext, useEffect, useState } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { usePnPService } from '../../../../common/services/usePnPService';
import { IFrameDialog, Pagination } from '@pnp/spfx-controls-react';
import { SearchQueryGeneratorForDocs, GraphSearchResponseMapper, DrivePermissionResponseMapper } from '../../../../common/utils/Utils';
import { SearchRequest, SearchResponse } from '@microsoft/microsoft-graph-types';
import { _CONST } from '../../../../common/utils/Const';
import { useGraphService } from '../../../../common/services/useGraphService';
import { ISearchResultExtended } from '../SharingView/ISearchResultExtended';
import { set } from '@microsoft/sp-lodash-subset';
import { IDrivePermissionParams } from '../../../../common/model';
import { ISharingResult } from '../../model';
import GoverFilter from '../helper/GoverFilter';
import { Toolbar } from '@pnp/spfx-controls-react/lib/Toolbar';

const SharingDetailedList: React.FC = (): JSX.Element => {

    const governContext = useContext(SharingWebPartContext);

    const { getSiteGroups
    } = usePnPService(governContext.webpartContext);
    const { getByGraphSearch, getDriveItemsPermission } = useGraphService(governContext.webpartContext);

    const [sharedFiles, setSharedFiles] = useState<ISharingResult[]>([]);
    const [fileIds, setFileIds] = useState<string[]>([]);
    const [spGroups, setSpGroups] = useState<string[]>();

    const [hideSharingSettingsDialog, setHideSharingSettingsDialog] = useState<boolean>(true);
    const [selectedDocument, setSelectedDocument] = useState<ISharingResult>();

    // let searchItems: Record<string, any> = [];
    const [searchItems, setSearchItems] = useState<ISearchResultExtended[]>([]);
    const [currentPage, setCurrentPage] = useState<number>();
    const [totalPages, setTotalPages] = useState<number>(1);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string>("");

    const loadPage = async (paramFileIds: string[]): Promise<void> => {
        try {
            setLoading(true);
            // const paramFileIds = paramFileIds || fileIds;
            const lastIndex = currentPage * governContext.pageLimit;
            const firstIndex = lastIndex - governContext.pageLimit;
            const paginatedItems = paramFileIds.slice(firstIndex, lastIndex);

            if (paginatedItems.length === 0) {
                console.log("No items to display");
                setLoading(false);
                return;
            }

            const paginatedListItems = paginatedItems.reduce((acc, fileId) => {
                const foundItem = searchItems.filter(item => item.FileId === fileId);
                if (foundItem.length === 0) return null;
                acc[fileId] = {
                    driveId: foundItem[0].DriveId,
                    driveItemId: foundItem[0].DriveItemId
                };
                return acc;
            }, {} as Record<string, IDrivePermissionParams>);
            let locSpGroups = spGroups;
            if (locSpGroups === undefined) {
                locSpGroups = await getSiteGroups();
                setSpGroups(locSpGroups);
            }
            // const sharedLinkResults = await getSharingLinks(searchItems, locSpGroups);

            try {
                // get searchItems where fileIds are in paginatedItems
                const locSearchItems = searchItems.filter(item => paginatedItems.includes(item.FileId));
                const sharedResults: ISharingResult[] = [];
                // const driveItemParam = locSearchItems.map(item => ({ driveId: item.DriveId, driveItemId: item.DriveItemId }));
                const driveItems = await getDriveItemsPermission(paginatedListItems);
                console.log("FazLog ~ loadPage ~ driveItems:", driveItems);

                // now we have all the data we need, we can start building up the result
                driveItems.forEach(driveItem => {
                    const file = locSearchItems.filter(item => item.FileId === driveItem.fileId)[0];
                    const locSharedResult = DrivePermissionResponseMapper(file, driveItem, locSpGroups);
                    sharedResults.push(locSharedResult);
                });

                if (!sharedResults) {
                    setLoading(false);
                    return;
                }
                const sharingLinks = sharedResults.filter(result => result.SharedWith !== null);
                setSharedFiles(sharingLinks);
            }
            catch (error) {
                throw error;
            }
        } catch (error) {
            console.error("Error loading page:", error);
            setError("Error loading page");
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        if (currentPage !== undefined) {
            loadPage(fileIds);
        }
    }, [currentPage]);

    useEffect(() => {
        const init = async (): Promise<void> => {
            try {
                const searchReqForDocs: SearchRequest = {
                    entityTypes: _CONST.GraphSearch.DocsSearch.EntityType,
                    query: {
                        queryString: `${SearchQueryGeneratorForDocs(governContext.webpartContext)}`
                    },
                    fields: _CONST.GraphSearch.DocsSearch.Fields,
                    from: 0,
                    size: 500
                };

                const searchResponse = await getByGraphSearch(searchReqForDocs);
                console.log("FazLog ~ init ~ searchResponse:", searchResponse);
                const locSearchItems = GraphSearchResponseMapper<ISearchResultExtended>(searchResponse, _CONST.GraphSearch.DocsSearch.EntityType);
                console.log("FazLog ~ init ~ locSearchItems:", locSearchItems);
                setSearchItems(locSearchItems);
                // get all file ids
                const locFileIds = locSearchItems.map((item) => item.FileId);
                setFileIds(locFileIds);
                setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));
                setCurrentPage(1);
                // await loadPage(locFileIds);
            } catch (error) {
                console.error("Error initializing:", error);
            }
        };
        init();
    }, []);

    const columns = [
        {
            key: 'SharingUserType',
            name: 'SharingUserType',
            fieldName: 'SharingUserType',
            minWidth: 16,
            maxWidth: 16,
            isIconOnly: true,
            isResizable: false
        },
        {
            key: 'FileExtension',
            name: 'FileExtension',
            fieldName: 'FileExtension',
            minWidth: 16,
            maxWidth: 16,
            isIconOnly: true,
            isResizable: false
        },
        {
            key: 'FileName',
            name: 'Filename',
            fieldName: 'FileName',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
            //isSorted: true,
            isSortedDescending: false,
            isRowHeader: true,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string'
        },
        {
            key: 'SharedWith',
            name: 'Shared with',
            fieldName: 'SharedWith',
            minWidth: 150,
            maxWidth: 185,
            isResizable: true,
            onRender: (item: ISharingResult) => {
                if (item.SharedWith === null)
                    return <span />;
                if (item.SharedWith.length > 1) {
                    return <Facepile
                        personaSize={PersonaSize.size24}
                        maxDisplayablePersonas={5}
                        personas={item.SharedWith}
                        overflowButtonType={OverflowButtonType.descriptive}
                        overflowButtonProps={{
                            ariaLabel: 'More users'
                        }}
                        ariaDescription="List of people who has been shared with."
                        ariaLabel="List of people who has been shared with."
                    />
                }
                else {
                    switch (item.SharingUserType) {
                        case "Link": return <Persona text={`${item.SharedWith[0].personaName}`} hidePersonaDetails={true} size={PersonaSize.size24} />; break;
                        default:
                            return <Persona text={`${item.SharedWith[0].personaName}`} hidePersonaDetails={true} size={PersonaSize.size24} />; break;
                    }
                }
            },
        },
        {
            key: 'Channel',
            name: 'Channel / Folder',
            fieldName: 'Channel',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            data: 'string'
        },
        {
            key: 'LastModified',
            name: 'Last modified',
            fieldName: 'LastModified',
            minWidth: 70,
            maxWidth: 90,
            isResizable: true,
            isPadded: true,
            data: 'date'
        },
        {
            key: 'SiteUrl',
            name: 'Site',
            fieldName: 'SiteUrl',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            data: 'string'
        },
    ];

    const _renderItemColumn = (item: ISharingResult, index: number, column: IColumn): any => {
        const fieldContent = item[column.fieldName as keyof ISharingResult] as string;

        // in here we're going to make the column render differently based on the column name
        switch (column.key) {
            case 'FileExtension':
                switch (item.FileExtension) {
                    case "folder": return <Icon {...getFileTypeIconProps({ type: FileIconType.documentsFolder, size: 16, imageFileType: 'png' })} />; break;
                    default: return <Icon {...getFileTypeIconProps({ extension: `${item.FileExtension}`, size: 16, imageFileType: 'png' })} />; break;
                }
            case 'SharingUserType':
                switch (item.SharingUserType) {
                    case "Guest": return <TooltipHost content="Shared with guest/external users" id="guestTip">
                        <Icon aria-label="SecurityGroup" aria-describedby="guestTip" iconName="SecurityGroup" id="Guest" />
                    </TooltipHost>; break;
                    case "Everyone": return <TooltipHost content="Shared with everyone" id="everyoneTip">
                        <Icon aria-label="Family" aria-describedby="everyoneTip" iconName="Family" id="Family" />
                    </TooltipHost>; break;
                    case "Member": return <span />;
                    case "Link": return <TooltipHost content="Shared with organization" id="everyoneTip">
                        <Icon aria-label="Family" aria-describedby="everyoneTip" iconName="Family" id="Family" />
                    </TooltipHost>; break;
                    case "Inherited": return <TooltipHost content="Shared by inheritance" id="inheritedTip">
                        <Icon aria-label="PartyLeader" aria-describedby="inheritedTip" iconName="PartyLeader" id="PartyLeader" />
                    </TooltipHost>; break;
                }
                break;
            case 'LastModified':
                return <span>{moment(item.LastModified).format('LL')}</span>; break;
            case 'FileName':
                return <span><Text><Link href={`${item.Url}`}>{`${item.FileName}`}</Link></Text></span>; break;
            case 'Channel':
                return <span><Text><Link href={`${item.FolderUrl}`}>{`${item.Channel}`}</Link></Text></span>; break;
            default:
                return <span>{fieldContent}</span>; break;
        }
    }

    const _loadSharingDialogDetails = (): void => {
        alert("test");
        console.log("FazLog ~ selectedDocument:", selectedDocument);
        setHideSharingSettingsDialog(false);
    }

    const _handleItemInvoked = (item: ISharingResult): void => {
        console.log("FazLog ~ item:", item);
        setSelectedDocument(item);

    }

    const _selection = new Selection({
        onSelectionChanged: () => 
    });


    return (
        <div>

            <GoverFilter />

            <Toolbar actionGroups={{
                'share': {
                    'share': {
                        title: 'Sharing Settings',
                        iconName: 'Share',
                        onClick: () => _loadSharingDialogDetails()
                    }
                }
            }}
                filters={[
                    {
                        id: "f1",
                        title: "Guest/External Users",
                    }]}
                // onSelectedFiltersChange={this._onSelectedFiltersChange.bind(this)}
                find={true}
            // onFindQueryChange={this._findItem}
            />
            {selectedDocument &&
                <IFrameDialog
                    url={`${selectedDocument.SiteUrl}/_layouts/15/sharedialog.aspx?listId=${selectedDocument.ListId}&listItemId=${selectedDocument.ListItemId}&clientId=sharePoint&mode=manageAccess&ma=0`}
                    // iframeOnLoad={this._onIframeLoaded.bind(this)}
                    hidden={hideSharingSettingsDialog}
                    onDismiss={() => setHideSharingSettingsDialog(true)}
                    modalProps={{
                        isBlocking: false
                    }}
                    dialogContentProps={{
                        type: DialogType.close,
                        showCloseButton: false
                    }}
                    width={'570px'}
                    height={'815px'}
                />
            }


            <MarqueeSelection selection={_selection}>
                <ShimmeredDetailsList
                    // usePageCache={true}
                    columns={columns}
                    // enableShimmer={(!loading)}
                    items={sharedFiles}
                    // selection={this.selection}
                    // onItemInvoked={this._handleItemInvoked}
                    selectionMode={SelectionMode.single}
                    onRenderItemColumn={_renderItemColumn}
                    onItemInvoked={(item) => _handleItemInvoked(item)}
                />
            </MarqueeSelection>
            <Pagination
                key="files"
                currentPage={currentPage}
                totalPages={totalPages}
                onChange={async (page) => {
                    setCurrentPage(page);
                    // await loadPage(page);
                }}
                limiter={3} // Optional - default value 3
                hideFirstPageJump // Optional
                hideLastPageJump // Optional
                limiterIcon={"Emoji12"} // Optional
            />
        </div>
    )
}

export default SharingDetailedList;