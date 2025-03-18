

import * as React from 'react'
import { TooltipHost, Icon, PersonaSize, Persona, Text, Link } from '@fluentui/react';
import { ISharedUser, SharedType } from '../../../../../common/model';
import { useEffect, useRef, useState } from 'react';

interface ISharedWithColumnProps {
    sharedType: SharedType;
    sharedWith: ISharedUser[];
    filteredSharedTypes: SharedType[];
}


const SharedWithColumn: React.FC<ISharedWithColumnProps> = ({ sharedWith, sharedType, filteredSharedTypes }): JSX.Element => {
    const [isExpanded, setIsExpanded] = useState(false);
    const contentRef = useRef<HTMLDivElement | null>(null);
    const [showButton, setShowButton] = useState(false);

    useEffect(() => {
        if (contentRef.current && contentRef.current.scrollHeight > 90) {
            setShowButton(true);
        }
    }, []);

    let locSharedWith = sharedWith;
    if (filteredSharedTypes.length > 0) {
        locSharedWith = locSharedWith.filter(val => {
            if (filteredSharedTypes.indexOf(val.type) > -1) {
                return val;
            }
        })
    }
    const handleShowMore = (): void => {
        setIsExpanded(true);
        setShowButton(false);
    };
    return (
        <>
            <div ref={contentRef}
                className="content"
                style={{
                    maxHeight: isExpanded ? 'none' : '90px',
                    overflow: 'hidden',
                    transition: 'max-height 0.3s ease',
                }}>
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
                    return <div style={{ paddingBottom: 4 }} key={sharedMember.id}>
                        <Persona
                            size={PersonaSize.size24}
                            imageAlt={sharedMember?.displayName || ''}
                            imageUrl={`${window.location.origin}/_layouts/15/userphoto.aspx?size=M&username=${sharedMember?.id}`}
                            text={sharedMember?.displayName || ''}
                            secondaryText={sharedMember?.id}
                        />
                    </div>
                })}
            </div>
            {showButton && !isExpanded && (
                <Link onClick={handleShowMore}>Show More</Link>
            )}

        </>
    );


}

export default SharedWithColumn;