
import * as React from 'react'
import { Identity, ItemActionSet, ItemActivityOLD, NullableOption } from '@microsoft/microsoft-graph-types-beta';
import { ActivityItem, IActivityItemProps, Link, mergeStyleSets, Persona, PersonaSize, TooltipHost } from '@fluentui/react';
import * as moment from 'moment';

interface IItemActivitiesProps {
    items: ItemActivityOLD[];
}


const getItemAction = (value: NullableOption<ItemActionSet> | undefined): JSX.Element | undefined => {
    if (value?.comment) {
        return <span><b>commented</b></span>;
    }
    if (value?.create) {
        return <span><b>created</b></span>;
    }
    if (value?.delete) {
        return <span><b>deleted</b></span>;
    }
    if (value?.edit) {
        return <span><b>edited</b></span>;
    }
    if (value?.mention) {
        return <span><b>mentioned</b> {value.mention.mentionees?.map((val, index) => <Persona
            key={index}
            text={val.user?.displayName || ""}
            size={PersonaSize.size8}
            imageAlt={val.user?.displayName || ''}
        />)}</span>;
    }
    if (value?.move) {
        const move = value.move;
        if (move.from && move.to) {
            return <span><b>moved</b> from {move.from} to {move.to}</span>
        } else if (move.from) {
            return <span><b>moved</b> from {move.from}</span>
        } else if (move.to) {
            return <span><b>moved</b> to {move.to}</span>;
        }
        return <span><b>moved</b></span>;
    }
    if (value?.rename) {
        if (value.rename?.newName && value.rename?.oldName) {
            return <span><b>renamed</b> from {value.rename.oldName} to {value.rename.newName}</span>;
        } else if (value.rename?.newName) {
            return <span><b>renamed</b> to {value.rename.newName}</span>;
        } else if (value.rename?.oldName) {
            return <span><b>renamed</b> from {value.rename.oldName}</span>;
        }
        return <span><b>renamed</b></span>;
    }
    if (value?.restore) {
        return <span><b>restored</b></span>;
    }
    if (value?.share) {
        if (value.share.recipients) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return <span><b>shared</b> with {value.share.recipients.map((val: any, index) => <Persona
                key={index}
                text={val.user?.displayName || val.user?.email}
                size={PersonaSize.size8}
                imageAlt={val.user?.displayName || ''}
            />)}</span>;
        }
        return <span><b>shared</b></span>;
    }
    if (value?.version) {
        return <span><b>version</b> {value.version?.newVersion}</span>;
    }
    return undefined;
}


const ItemActivities: React.FC<IItemActivitiesProps> = ({ items }): JSX.Element => {
    const classNames = mergeStyleSets({
        activityRoot: {
            marginTop: '20px',
        },
        nameText: {
            fontWeight: 'bold',
        },
    });

    const activityItems: (IActivityItemProps & { key: string | number })[] = items.map((item, index) => {
        const actor = item.actor?.user as Identity & { email: string };
        return {
            key: index,
            activityDescription: [
                <div key={index} id={index + "0"} style={{ display: "flex" }}>
                    <Link
                        key={index + "1"}
                    >
                        {/* {actor?.displayName || ''} */}
                        <TooltipHost content={actor?.email || ""} id='actor'>
                            <Persona
                                text={item.actor?.user?.displayName || ""}
                                size={PersonaSize.size8}
                                imageAlt={item.actor?.user?.displayName || ''}
                                imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${actor?.email}`}
                            />
                        </TooltipHost>
                    </Link>
                    <div key={index + "2"}> {getItemAction(item.action) ?? ""} </div>
                </div>
            ],
            // activityDescription: [
            //         {/* <TooltipHost content={actor?.email || ""} id='actor'> */}
            //         // <Persona
            //         //     text={item.actor?.user?.displayName || ""}
            //         //     size={PersonaSize.size8}
            //         //     imageAlt={item.actor?.user?.displayName || ''}
            //         //     imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${actor?.email}`}
            //         // />
            //         {/* </TooltipHost> */}
            //         <Link
            //             key={1}
            //             className={classNames.nameText}
            //             onClick={() => {
            //                 alert('A name was clicked.');
            //             }}
            //         >
            //             {item.actor?.user?.displayName || ''}
            //         </Link>,
            //         <span key={index + "2"}> {getItemAction(item.action) ?? ""} </span>
            // ],
            timeStamp: moment(item.times?.recordedDateTime).format('DD/MM/YY HH:mm')
        }
    });


    return <div>
        {activityItems.map((item: { key: string | number }) => (
            <ActivityItem {...item} key={item.key} className={classNames.activityRoot} />
        ))}
    </div>
}

export default ItemActivities;