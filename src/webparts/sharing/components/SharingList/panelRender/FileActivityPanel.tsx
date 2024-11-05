

import * as React from 'react'
import { IFileSharingResponse } from '../../../../../common/model';
import { useGraphService } from '../../../../../common/services/useGraphService';
import { SharingWebPartContext } from '../../../hooks/SharingWebPartContext';
import { useEffect, useState } from 'react';
import { ItemActivityOLD } from '@microsoft/microsoft-graph-types-beta';
import ItemActivities from '../../helper/ItemActivities';
import { MessageBar, MessageBarType, Spinner } from '@fluentui/react';

interface IFileActivityPanelProps {
    file: IFileSharingResponse;
}


const FileActivityPanel: React.FC<IFileActivityPanelProps> = ({ file }): JSX.Element => {
    const governContext = React.useContext(SharingWebPartContext);
    const webpartContext = governContext.webpartContext;
    const { getItemsActivityBETA } = useGraphService(webpartContext);

    const [itemActivities, setItemActivities] = useState<ItemActivityOLD[] | undefined>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string>("");

    useEffect(() => {
        const getActivity = async (): Promise<void> => {
            try {
                setError("");
                const responseItemActivity = await getItemsActivityBETA({
                    driveId: file.DriveId,
                    itemId: file.FileId
                });
                setItemActivities(responseItemActivity);
            } catch (error) {
                setError(error);
            }
            setLoading(false);
        }
        getActivity().catch(e => console.error(e));
    }, []);

    if (loading) {
        return <div style={{ paddingTop: "50%" }}>
            <Spinner label='loading activities...' />
        </div>
    }

    if (error) {
        return <MessageBar messageBarType={MessageBarType.error}>
            Something went wrong while fetching activities.
        </MessageBar>
    }

    return <div>
        {itemActivities?.length === 0 && <div>
            <MessageBar messageBarType={MessageBarType.info}>
                No activities found.
            </MessageBar>
        </div>}

        {itemActivities && itemActivities.length > 0 && <div>
            <ItemActivities items={itemActivities} />
        </div>}
    </div>
}

export default FileActivityPanel;