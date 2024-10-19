/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */


import { Facepile, IColumn, Icon, Link, OverflowButtonType, Persona, Text, PersonaSize, ShimmeredDetailsList, TooltipHost, SelectionMode } from '@fluentui/react';
import * as React from 'react'
import * as moment from 'moment';
import { FileIconType, getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import ISharingResult from '../SharingView/ISharingResult';
import { useContext, useEffect, useState } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { useDataProvider } from '../../../../services/useDataProvider';
import { Pagination } from '@pnp/spfx-controls-react';

const SharingDetailedList: React.FC = (): JSX.Element => {

    const governContext = useContext(SharingWebPartContext);
    const { loadAssociatedGroups, getSharingLinks, getSearchResults } = useDataProvider();

    const [sharedFiles, setSharedFiles] = useState<ISharingResult[]>([]);
    const [fileIds, setFileIds] = useState<string[]>([]);
    // const [searchItems, setSearchItems] = useState<Record<string, any>>([]);
    let searchItems: Record<string, any> = [];
    const [currentPage, setCurrentPage] = useState<number>(1);
    const [totalPages, setTotalPages] = useState<number>(1);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string>("");

    const _processSharingLinks = async (locFieldIds: string[]): Promise<ISharingResult[]> => {
        try {
            // setting these to be empty because of the pagination, 
            // otherwise in the pagination items will be added to the existing array
            const sharingLinks: ISharingResult[] = [];

            const paginatedListItems: Record<string, any> = {};
            locFieldIds.forEach((fileId) => {
                paginatedListItems[fileId] = searchItems[fileId];
            });


            // getting the sharing links using expensive REST API calls based on the given list of Id's 
            console.log("FazLog ~ const_processSharingLinks= ~ paginatedListItems:", paginatedListItems);
            const sharedLinkResults = await getSharingLinks(paginatedListItems);
            if (sharedLinkResults === null)
                return;

            sharedLinkResults.forEach((sharedResult) => {
                if (sharedResult.SharedWith === null)
                    return;

                sharingLinks.push(sharedResult)
            });

            return (sharingLinks);
        } catch (error) {
            console.log("FazLog ~ error:", error);
        }
    }


    const loadPage = async (page: number, paramFileIds?: string[]): Promise<void> => {
        try {
            const locFileIds = paramFileIds ? paramFileIds : fileIds;
            setLoading(true);
            const lastIndex = page * governContext.pageLimit;
            const firstIndex = lastIndex - governContext.pageLimit;

            const paginatedItems = locFileIds.slice(firstIndex, lastIndex);
            setCurrentPage(page);
            if (paginatedItems.length === 0) {
                console.log("No items to display");
                return;
            }
            else {
                console.log(`${locFileIds.length} shared items found`);
            }


            const locSharedFiles: ISharingResult[] = await _processSharingLinks(paginatedItems);
            setSharedFiles(locSharedFiles);


        } catch (error) {
            console.log("FazLog ~ loadPage= ~ error:", error);
            setError("In loadPage");
        }
        setLoading(false);
    }

    useEffect(() => {
        const init = async (): Promise<void> => {
            await loadAssociatedGroups();

            try {
                const locSearchItems = await getSearchResults();
                console.log("FazLog ~ init ~ locSearchItems:", locSearchItems);
                searchItems = locSearchItems;
                // setSearchItems(locSearchItems);
                const locFileIds = Object.keys(locSearchItems);
                setFileIds(locFileIds);
                setTotalPages(Math.ceil(locFileIds.length / governContext.pageLimit));
                await loadPage(currentPage, locFileIds);

            } catch (error) {
                console.log("FazLog ~ init ~ error:", error);
            }
        };
        init().catch((error) => console.error(error));
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


    return (
        <div>

            <ShimmeredDetailsList
                usePageCache={true}
                columns={columns}
                enableShimmer={(!loading)}
                items={sharedFiles}
                // selection={this.selection}
                // onItemInvoked={this._handleItemInvoked}
                selectionMode={SelectionMode.single}
                onRenderItemColumn={_renderItemColumn}
            />
            <Pagination
                key="files"
                currentPage={currentPage}
                totalPages={totalPages}
                onChange={async (page) => {
                    await loadPage(page);
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