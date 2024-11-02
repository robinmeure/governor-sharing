

import * as React from 'react'
import { IFileSharingResponse } from '../../../../../common/model';
import { useGraphService } from '../../../../../common/services/useGraphService';
import { SharingWebPartContext } from '../../../hooks/SharingWebPartContext';
import { useEffect, useState } from 'react';
import { ItemActivityOLD } from '@microsoft/microsoft-graph-types-beta';
import ItemActivities from '../../helper/ItemActivities';

interface IFileActivityPanelProps {
    file: IFileSharingResponse;
}


const FileActivityPanel: React.FC<IFileActivityPanelProps> = ({ file }): JSX.Element => {
    const governContext = React.useContext(SharingWebPartContext);
    const webpartContext = governContext.webpartContext;
    const { getItemsActivity } = useGraphService(webpartContext);

    const [itemActivities, setItemActivities] = useState<ItemActivityOLD[] | undefined>([]);

    useEffect(() => {
        const getActivity = async (): Promise<void> => {
            const responseItemActivity = await getItemsActivity({
                driveId: file.DriveId,
                itemId: file.FileId
            });
            console.log("FazLog ~ getActivity ~ responseItemActivity:", responseItemActivity);
            setItemActivities(responseItemActivity);
        }
        getActivity().catch(e => console.error(e));
    }, []);

    return <div>
        <div>
            {itemActivities && <div>
                {itemActivities.map(itemAct => {
                    return <div key={itemAct.id}>
                        <ItemActivities item={itemAct} />
                    </div>
                })}
                Activities
                {/* {itemActivity.map(val => {
                    return <div key={val.id}>
                        {val.id}
                        {val.action}
                        {val.actor}
                        {val.times}
                    </div>
                })} */}
            </div>}
        </div>
    </div>
}

export default FileActivityPanel;