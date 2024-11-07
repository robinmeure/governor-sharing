
import * as React from 'react'
import { ItemActionSet, ItemActivityOLD, NullableOption } from '@microsoft/microsoft-graph-types-beta';
import { Persona, PersonaSize, Text } from '@fluentui/react';
import * as moment from 'moment';

interface IItemActivitiesProps {
    items: ItemActivityOLD[];
}


const getItemAction = (value: NullableOption<ItemActionSet> | undefined): JSX.Element | undefined => {
    if (value?.comment) {
        return <span> <i>commented</i></span>;
    }
    if (value?.create) {
        return <span> <i>created</i></span>;
    }
    if (value?.delete) {
        return <span> <i>deleted</i></span>;
    }
    if (value?.edit) {
        return <span> <i>edited</i></span>;
    }
    if (value?.mention) {
        return <span><i>mentioned</i> {value.mention.mentionees?.map((val, index) => <Persona
            key={index}
            text={val.user?.displayName || ""}
            size={PersonaSize.size8}
            imageAlt={val.user?.displayName || ''}
        />)}</span>;
    }
    if (value?.move) {
        const move = value.move;
        if (move.from && move.to) {
            return <span><i>moved</i> from {move.from} to {move.to}</span>
        } else if (move.from) {
            return <span><i>moved</i> from {move.from}</span>
        } else if (move.to) {
            return <span><i>moved</i> to {move.to}</span>;
        }
        return <span><i>moved</i></span>;
    }
    if (value?.rename) {
        if (value.rename?.newName && value.rename?.oldName) {
            return <span><i>renamed</i> from {value.rename.oldName} to {value.rename.newName}</span>;
        } else if (value.rename?.newName) {
            return <span><i>renamed</i> to {value.rename.newName}</span>;
        } else if (value.rename?.oldName) {
            return <span><i>renamed</i> from {value.rename.oldName}</span>;
        }
        return <span><i>renamed</i></span>;
    }
    if (value?.restore) {
        return <span><i>restored</i></span>;
    }
    if (value?.share) {
        if (value.share.recipients) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return <span><i>shared</i> with {value.share.recipients.map((val: any, index) => (
                <span key={index}>
                    <Text variant="smallPlus">
                        <i>{val.user?.displayName || val.user?.email}</i>
                    </Text>
                </span>
            ))}</span>;
        }
        return <span><i>shared</i></span>;
    }
    if (value?.version) {
        return <span><i>version</i> {value.version?.newVersion}</span>;
    }
    return undefined;
}


const ItemActivities: React.FC<IItemActivitiesProps> = ({ items }): JSX.Element => {
    // const classNames = mergeStyleSets({
    //     activityRoot: {
    //         marginTop: '20px',
    //     },
    //     nameText: {
    //         fontWeight: 'bold',
    //     },
    // });

    // const activityItems: (IActivityItemProps & { key: string | number })[] = items.map((item, index) => {
    //     const actor = item.actor?.user as Identity & { email: string };
    //     return {
    //         key: index,
    //         activityDescription: [
    //             <div key={index} id={index + "0"} style={{ display: "flex" }}>

    //                 <TooltipHost content={actor?.email || ""} id='actor'>
    //                     <Persona
    //                         text={item.actor?.user?.displayName || ""}
    //                         size={PersonaSize.size8}
    //                         imageAlt={item.actor?.user?.displayName || ''}
    //                         imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${actor?.email}`}
    //                     />
    //                 </TooltipHost>
    //                 <div key={index + "2"}> {getItemAction(item.action) ?? ""} </div>
    //             </div>
    //         ],
    //         timeStamp: moment(item.times?.recordedDateTime).format('DD/MM/YY HH:mm')
    //     }
    // });


    return <div>
        {items.map((item, index) => {
            return <div key={index} style={{ display: "flex", padding: "12px 6px", borderRadius: 6, backgroundColor: index % 2 === 1 ? '#f9f9f9' : 'transparent' }}>
                <Persona
                    text={item.actor?.user?.displayName || ""}
                    size={PersonaSize.size8}
                    imageAlt={item.actor?.user?.displayName || ''}
                    styles={{
                        details: {
                            paddingRight: 4,
                        }
                    }}
                />
                {getItemAction(item.action) ?? ""}
                &nbsp;on&nbsp;

                <Text variant="smallPlus" style={{ marginTop: 2 }}>
                    <i>
                        {moment(item.times?.recordedDateTime).format('DD/MM/YY HH:mm')}
                    </i>
                </Text>

            </div>

        })}
    </div>
}

export default ItemActivities;