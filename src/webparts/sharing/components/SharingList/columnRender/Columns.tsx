import { IColumn, Text, Persona, PersonaSize, ActionButton } from "@fluentui/react";
import * as moment from "moment";
import * as React from "react";
import { IFileSharingResponse } from "../../../../../common/model";
import { IPaginationFilterState } from "../SharingDetailedList";
import FileExtentionColumn from "./FileExtentionColumn";
import LinkColumn from "./LinkColumn";
import SharedWithColumn from "./SharedWithColumn";


interface IColumnsProps {
    paginationFilterState: IPaginationFilterState;
    setPaginationFilterState: React.Dispatch<React.SetStateAction<IPaginationFilterState>>;
}


export default function Columns({ paginationFilterState, setPaginationFilterState }: IColumnsProps): IColumn[] {
    // Define columns for the details list
    const columns: IColumn[] = [
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
    ];

    return columns;
}
