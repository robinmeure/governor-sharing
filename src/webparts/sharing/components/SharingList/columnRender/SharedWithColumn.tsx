

import * as React from 'react'
import { TooltipHost, Icon, IFacepilePersona, PersonaSize, Facepile, OverflowButtonType, Persona, Text } from '@fluentui/react';
import { SharedType } from '../../../../../common/model';

interface ISharedWithColumnProps {
    sharedType: SharedType;
    sharedWith: IFacepilePersona[];
}


const SharedWithColumn: React.FC<ISharedWithColumnProps> = ({ sharedWith, sharedType }): JSX.Element => {

    return (
        <>
            <div>

                {sharedWith?.length === 0 && <Persona text={`${sharedWith[0]?.personaName}`} hidePersonaDetails={true} size={PersonaSize.size24} />}

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
                )}

            </div>

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
                        <div style={{ marginTop: 4 }}>
                            <TooltipHost content={textContent} id="guestTip">
                                <Icon aria-label={iconName} iconName={iconName} id={iconName + "icon"} />
                                <Text style={{ marginLeft: 8 }} variant="small">{textContent}</Text>
                            </TooltipHost>
                        </div>

                    );
                })()}

            </div>
        </>
    );


}

export default SharedWithColumn;