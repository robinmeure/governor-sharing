
import * as React from 'react'
import { Checkbox, DefaultButton, Dropdown, IDropdownOption, Label, Panel, PrimaryButton, TextField } from '@fluentui/react';
import { useContext, useEffect, useState } from 'react';
import { ISiteSearchResponse, SharedType } from '../../../../../common/model';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { GraphSearchResponseMapper } from '../../../../../common/config/Mapper';
import { _CONST } from '../../../../../common/utils/Const';
import { useGraphService } from '../../../../../common/services/useGraphService';
import { SharingWebPartContext } from '../../../hooks/SharingWebPartContext';


export interface IFilterItem {
    siteUrl: string;
    modifiedBy: string;
    sharedType: SharedType[];

}

interface IFilterPanelProps {
    filterItem: IFilterItem;
    isFilterPanelOpen: boolean;
    onDismissFilterPanel(newFilterVal?: IFilterItem): void;
}


const FilterPanel: React.FC<IFilterPanelProps> = (props): JSX.Element => {

    const { webpartContext } = useContext(SharingWebPartContext);
    const { getByGraphSearch } = useGraphService(webpartContext);
    const [filtreVal, setFilterVal] = useState<IFilterItem>(props.filterItem);

    const [siteFilterOptions, setSiteFilterOptions] = useState<IDropdownOption[]>([]);

    useEffect(() => {
        const getFilerValues = async (): Promise<void> => {
            try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const searchReqForSites: SearchRequest | {} = {
                    entityTypes: _CONST.GraphSearch.SiteSearch.EntityType,
                    query: {
                        queryString: "*"
                    },
                    trimDuplicates: true // https://github.com/microsoftgraph/msgraph-sdk-java/issues/1315
                };
                const siteSearchRes = await getByGraphSearch(searchReqForSites);
                console.log("FazLog ~ init ~ siteSearchRes:", siteSearchRes);
                const locSearchItems = GraphSearchResponseMapper<ISiteSearchResponse>(siteSearchRes, _CONST.GraphSearch.SiteSearch.EntityType);
                console.log("FazLog ~ init ~ locSearchItems:", locSearchItems);
                const siteOptions = locSearchItems.map(item => {
                    const parsedUrl = new URL(item.url);
                    return {
                        key: item.url,
                        text: item.name + `(${parsedUrl.pathname})`
                    }
                });
                setSiteFilterOptions(siteOptions);

            } catch (error) {
                console.log("FazLog ~ getFilerValues ~ error:", error);
            }
        }
        getFilerValues().catch(error => console.log("FazLog ~ init ~ error", error));
    }, []);

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

                <div style={{ padding: "12px 0" }}>
                    <Label>Site Url</Label>
                    <div>

                        {siteFilterOptions?.length < 50 ? <>
                            <Dropdown
                                selectedKey={filtreVal.siteUrl}
                                options={siteFilterOptions} onChange={(e, op) => {
                                    setFilterVal({ ...filtreVal, siteUrl: op?.key as string })
                                }} />
                        </> : <>
                            <TextField
                                multiline
                                resizable={false}
                                value={filtreVal.siteUrl}
                                onChange={(e, newVal = '') => setFilterVal({ ...filtreVal, siteUrl: newVal })}
                                placeholder={`https://contoso.sharepoint.com/sites/...`}
                                description="Url of the site" />
                        </>}
                    </div>
                </div>

                <div style={{ padding: "12px 0" }}>
                    <Label>Shared Type</Label>
                    <div style={{ gap: "8px", display: "flex", flexDirection: "column" }}>
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
                        <Checkbox label="Groups"
                            checked={filtreVal.sharedType.filter(val => val === "Groups").length > 0}
                            onChange={(ex, checked) => {
                                console.log("FazLog ~ ex:", ex);
                                setFilterVal({
                                    ...filtreVal,
                                    sharedType: checked ? [...filtreVal.sharedType, "Groups"] : filtreVal.sharedType.filter(val => val !== "Groups")
                                })
                            }}
                        />
                        <Checkbox label="Member"
                            checked={filtreVal.sharedType.filter(val => val === "Member").length > 0}
                            onChange={(ex, checked) => {
                                console.log("FazLog ~ ex:", ex);
                                setFilterVal({
                                    ...filtreVal,
                                    sharedType: checked ? [...filtreVal.sharedType, "Member"] : filtreVal.sharedType.filter(val => val !== "Member")
                                })
                            }}
                        />
                    </div>

                </div>

            </div>
        </Panel>

    </div>
}

export default FilterPanel;