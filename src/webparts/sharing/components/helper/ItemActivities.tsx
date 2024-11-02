

import * as React from 'react'
import { ItemActivityOLD } from '@microsoft/microsoft-graph-types-beta';
import { ActivityItem, IActivityItemProps, mergeStyleSets } from '@fluentui/react';

interface IItemActivitiesProps {
    item: ItemActivityOLD;
}


const ItemActivities: React.FC<IItemActivitiesProps> = ({ item }): JSX.Element => {
    const classNames = mergeStyleSets({
        activityRoot: {
            marginTop: '20px',
        },
        nameText: {
            fontWeight: 'bold',
        },
    });

    const activityItems: (IActivityItemProps & { key: string | number })[] = [
        {
            key: 1,
            activityDescription: [
                // <Persona
                //     key={key + "1"}
                //     size={PersonaSize.size8}
                //     imageAlt="Annie Lindqvist, no presence detected"
                // />,
                <span key={2}> renamed </span>,
                <span key={3} className={classNames.nameText}>
                    {item.id}
                </span>,
            ],
            // activityPersonas: [{ imageUrl: TestImages.personaMale }],
            comments: 'Hello, this  !',
            timeStamp: '23m ago',
        }
    ];


    return <div>
        {activityItems.map((item: { key: string | number }) => (
            <ActivityItem {...item} key={item.key} className={classNames.activityRoot} />
        ))}
    </div>
}

export default ItemActivities;