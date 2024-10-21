/* eslint-disable */

import * as React from 'react';
import { IContextualMenuProps } from '@fluentui/react';
import { ISharingResult } from '@pnp/sp/sharing';
import SharingDetailedList from './SharingDetailedList';

export interface ISharingListState {
    files: ISharingResult[];
    fileIds: string[];
    searchItems: Record<string, any>;
    groups: any[];
    isOpen: boolean;
    selectedDocument: ISharingResult;
    hideSharingSettingsDialog: boolean;
    frameUrlSharingSettings: string;
    contextualMenuProps?: IContextualMenuProps;
    showResetFilters?: boolean;
    currentPage: number,
    totalPages: number,
    pageLimit: number,
    selectedFilter?: string,
    loadingComplete: boolean,
    statusMessage: string,

}

const SharingList: React.FC = (): JSX.Element => {


    return <>
        <SharingDetailedList />
    </>;
};

export default SharingList;



