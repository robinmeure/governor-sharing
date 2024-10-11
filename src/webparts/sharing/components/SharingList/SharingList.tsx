/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import { useContext, useEffect } from 'react';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { IContextualMenuProps } from '@fluentui/react';
import { ISharingResult } from '@pnp/sp/sharing';
import { useDataProvider } from '../../../../services/useDataProvider';

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

    const { loadAssociatedGroups } = useDataProvider(usdd.webpartContext);

    useEffect(() => {
        const init = async (): Promise<void> => {
            await loadAssociatedGroups();
            setTimeout(async () => {
                await loadAssociatedGroups();
            }, 3000);

            setTimeout(async () => {
                await loadAssociatedGroups("https://res4devconsultinginc.sharepoint.com/sites/ArchivedProjects");
            }, 5000);
        };
        init().catch((error) => console.error(error));
    }, []);

    return <>
        Another Test Sharing list
    </>;
};

export default SharingList;



