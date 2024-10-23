/* eslint-disable */

import { Dropdown, IDropdownOption, Stack } from '@fluentui/react';
import * as React from 'react';
import { useContext, useEffect, useState } from 'react';
import { usePnPService } from '../../../../common/services/usePnPService';
import { SharingWebPartContext } from '../../hooks/SharingWebPartContext';
import { useGraphService } from '../../../../common/services/useGraphService';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { GraphSearchResponseMapper } from '../../../../common/utils/Utils';
import { ISiteData } from '../../../../common/model';
import { _CONST } from '../../../../common/utils/Const';


const GoverFilter: React.FC = (): JSX.Element => {

    const governContext = useContext(SharingWebPartContext);
    const { getByGraphSearch } = useGraphService(governContext.webpartContext);

    const [siteFilter, setSiteFilter] = useState<string[]>([]);
    const [libFilter, setLibFilter] = useState<string[]>([]);

    const options: IDropdownOption[] = [

    ];

    useEffect(() => {
        const getFilerValues = async () => {
            try {
                const searchReqForSites: SearchRequest = {
                    entityTypes: _CONST.GraphSearch.SiteSearch.EntityType,
                    query: {
                        queryString: "*"
                    }
                };
                const siteSearchRes = await getByGraphSearch(searchReqForSites);
                console.log("FazLog ~ init ~ siteSearchRes:", siteSearchRes);
                const locSearchItems = GraphSearchResponseMapper<ISiteData>(siteSearchRes, _CONST.GraphSearch.SiteSearch.EntityType);
                console.log("FazLog ~ init ~ locSearchItems:", locSearchItems);

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



