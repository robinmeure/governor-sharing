/* eslint-disable */

import { Dropdown, IDropdownOption, Stack } from '@fluentui/react';
import * as React from 'react';
import { useContext, useEffect, useState } from 'react';
import { usePnPService } from '../../../../common/services/usePnPService';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';


const GoverFilter: React.FC = (): JSX.Element => {

    const governContext = useContext(SharingWebPartContext);
    const { getSiteGroups,
        // getSharingLinks, 
        // getSearchResults,
    } = usePnPService(governContext.webpartContext);
    const [siteFilter, setSiteFilter] = useState<string[]>([]);
    const [libFilter, setLibFilter] = useState<string[]>([]);

    const options: IDropdownOption[] = [

    ];

    useEffect(() => {
        const getFilerValues = () => {
            try {

            } catch (error) {
                console.log("FazLog ~ getFilerValues ~ error:", error);
            }
        }
        getFilerValues();
    }, []);

    return <div>
        <Stack enableScopedSelectors horizontal horizontalAlign="space-between">
            <div>
                <Dropdown
                    placeholder="Select an option"
                    label="Basic uncontrolled example"
                    options={options}
                />
            </div>

            <div>
                <Dropdown
                    placeholder="Select an option"
                    label="Basic uncontrolled example"
                    options={options}
                />
            </div>
        </Stack>
    </div>;
};

export default GoverFilter;



