
import * as React from 'react'
import { Checkbox, DefaultButton, Label, Panel, PrimaryButton, TextField } from '@fluentui/react';
import { useState } from 'react';
import { SharedType } from '../../../../../common/model';


export interface IFilterItem {
    siteUrl: string;
    sharedType: SharedType[];
    modifiedBy: string;

}




interface IFilterPanelProps {
    filterItem: IFilterItem;
    isFilterPanelOpen: boolean;
    onDismissFilterPanel(newFilterVal?: IFilterItem): void;
}


const FilterPanel: React.FC<IFilterPanelProps> = (props): JSX.Element => {

    // const { getByGraphSearch } = useGraphService(governContext.webpartContext);
    const [filtreVal, setFilterVal] = useState<IFilterItem>(props.filterItem);

    // const [siteFilterOptions, setSiteFilterOptions] = useState<IDropdownOption[]>([]);

    // useEffect(() => {
    //     const getFilerValues = async (): Promise<void> => {
    //         try {
    //             const searchReqForSites: SearchRequest = {
    //                 entityTypes: _CONST.GraphSearch.SiteSearch.EntityType,
    //                 query: {
    //                     queryString: "*"
    //                 },
    //                 from: 0,
    //                 size: 500
    //             };
    //             const siteSearchRes = await getByGraphSearch(searchReqForSites);
    //             console.log("FazLog ~ init ~ siteSearchRes:", siteSearchRes);
    //             const locSearchItems = GraphSearchResponseMapper<ISiteSearchResponse>(siteSearchRes, _CONST.GraphSearch.SiteSearch.EntityType);
    //             console.log("FazLog ~ init ~ locSearchItems:", locSearchItems);
    //             const siteOptions = locSearchItems.map(item => {
    //                 return {
    //                     key: item.url,
    //                     text: item.name
    //                 }
    //             });
    //             setSiteFilterOptions(siteOptions);

    //         } catch (error) {
    //             console.log("FazLog ~ getFilerValues ~ error:", error);
    //         }
    //     }
    //     getFilerValues().catch(error => console.log("FazLog ~ init ~ error", error));
    // }, []);

    return <div>
        <Panel
            headerText="Filter"
            isOpen={props.isFilterPanelOpen}
            onDismiss={() => props.onDismissFilterPanel()}
            closeButtonAriaLabel="Close"
            onRenderFooterContent={() =>
                <div>
                    <PrimaryButton onClick={() => {
                        props.onDismissFilterPanel(filtreVal);
                    }}>Apply</PrimaryButton>
                    <DefaultButton style={{ marginLeft: 16 }} onClick={() => {
                        setFilterVal(props.filterItem);
                        props.onDismissFilterPanel();
                    }}>Cancel</DefaultButton>
                </div>}
            isFooterAtBottom={true}
        >
            <div>

                <div>
                    <Label>Site Url</Label>
                    <div>
                        <TextField
                            multiline
                            resizable={false}
                            value={filtreVal.siteUrl}
                            onChange={(e, newVal = '') => setFilterVal({ ...filtreVal, siteUrl: newVal })}
                            placeholder={`https://contoso.sharepoint.com/sites/...`}
                            description="Url of the site" />
                    </div>

                </div>

                <div>
                    <Label>Shared Type</Label>

                    <Checkbox label="Guest/External Users"
                        checked={filtreVal.sharedType.filter(val => val === "Guest").length > 0}
                        onChange={(ex, checked) => {
                            console.log("FazLog ~ ex:", ex);
                            setFilterVal({
                                ...filtreVal,
                                sharedType: checked ? [...filtreVal.sharedType, "Guest"] : filtreVal.sharedType.filter(val => val !== "Guest")
                            })
                        }}
                    />
                </div>

            </div>
        </Panel>

    </div>
}

export default FilterPanel;