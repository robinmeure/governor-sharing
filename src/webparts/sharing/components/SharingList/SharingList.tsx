/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import { useContext } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { IContextualMenuProps } from '@fluentui/react';
import { ISharingResult } from '@pnp/sp/sharing';

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

    const usdd = useContext(SharingWebPartContext);
    console.log("FazLog ~ usdd:", usdd.isTeams);

    return <>
        Sharing list
    </>;
};

export default SharingList;



