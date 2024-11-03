

import * as React from 'react'
import { TooltipHost, Icon, PersonaSize, Persona, Text } from '@fluentui/react';
import { ISharedUser, SharedType } from '../../../../../common/model';

interface ISharedWithColumnProps {
    sharedType: SharedType;
    sharedWith: ISharedUser[];
    filteredSharedTypes: SharedType[];
}


const SharedWithColumn: React.FC<ISharedWithColumnProps> = ({ sharedWith, sharedType, filteredSharedTypes }): JSX.Element => {


    let locSharedWith = sharedWith;
    if (filteredSharedTypes.length > 0) {
        locSharedWith = locSharedWith.filter(val => {
            if (filteredSharedTypes.indexOf(val.type) > -1) {
                return val;
            }
        })
    }
    return (
        <>
            <div>

                <div>
                    {(() => {
                        let textContent = "", iconName = "";
                        switch (sharedType) {
                            case "Guest":
                                textContent = "Shared with guest/external users";
                                iconName = "SecurityGroup";
                                break;
                            case "Everyone":
                                textContent = "Shared with everyone";
                                iconName = "Family";
                                break;
                            case "Link":
                                textContent = "Shared with organization";
                                iconName = "Family";
                                break;
                            case "Inherited":
                                textContent = "Shared by inheritance";
                                iconName = "PartyLeader";
                                break;
                            default:
                                return <></>;
                        }
                        return (
                            <div style={{ marginBottom: 4 }}>
                                <TooltipHost content={textContent} id="guestTip">
                                    <Icon aria-label={iconName} iconName={iconName} id={iconName + "icon"} />
                                    <Text style={{ marginLeft: 8 }} variant="small">{textContent}</Text>
                                </TooltipHost>
                            </div>

                        );
                    })()}

                </div>

                {locSharedWith.map((sharedMember) => {
                    return <div key={sharedMember.id}>
                        <Persona
                            size={PersonaSize.size24}
                            imageAlt={sharedMember?.displayName || ''}
                            imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${sharedMember?.id}`}
                            text={sharedMember?.displayName || ''}
                            secondaryText={sharedMember?.id}
                        />
                    </div>
                })}

                {/* {sharedWith?.length === 0 && <Persona text={`${sharedWith[0]?.name}`} hidePersonaDetails={true} size={PersonaSize.size24} />}

                {sharedWith?.length > 1 && (
                    <Facepile
                        personaSize={PersonaSize.size24}
                        maxDisplayablePersonas={5}
                        personas={sharedWith}
                        overflowButtonType={OverflowButtonType.descriptive}
                        overflowButtonProps={{
                            ariaLabel: 'More users'
                        }}
                        ariaDescription="List of people who has been shared with."
                        ariaLabel="List of people who has been shared with."
                    />
                )} */}

            </div>


        </>
    );


}

export default SharedWithColumn;