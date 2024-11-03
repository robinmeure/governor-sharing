
import * as React from 'react'
import { Checkbox, ChoiceGroup, DefaultButton, Dropdown, IChoiceGroupOption, IDropdownOption, Label, Panel, PrimaryButton, TextField } from '@fluentui/react';
import { useEffect, useState } from 'react';
import { ISiteSearchResponse, SharedType } from '../../../../../common/model';
import { SearchRequest } from '@microsoft/microsoft-graph-types';
import { GraphSearchResponseMapper } from '../../../../../common/config/Mapper';
import { useGraphService } from '../../../../../common/services/useGraphService';
import { _CONST } from '../../../../../common/utils/Const';
import { SharingWebPartContext } from '../../../hooks/SharingWebPartContext';

type FileFolderFilter = "OnlyFiles" | "OnlyFolders" | "BothFilesFolders";
export interface IFilterItem {
    siteUrl: string;
    modifiedBy: string;
    sharedType: SharedType[];
    fileFolder: FileFolderFilter;
}

interface IFilterPanelProps {
    filterItem: IFilterItem;
    isFilterPanelOpen: boolean;
    onDismissFilterPanel(newFilterVal?: IFilterItem): void;
}


const FilterPanel: React.FC<IFilterPanelProps> = (props): JSX.Element => {

    const { webpartContext } = React.useContext(SharingWebPartContext);
    const { getByGraphSearch } = useGraphService(webpartContext);
    const [filtreVal, setFilterVal] = useState<IFilterItem>(props.filterItem);

    const [siteFilterOptions, setSiteFilterOptions] = useState<IDropdownOption[]>([]);

    useEffect(() => {
        const getFilerValues = async (): Promise<void> => {
            const searchReqForSites: SearchRequest & { trimDuplicates: boolean } = {
                entityTypes: _CONST.GraphSearch.SiteSearch.EntityType,
                query: {
                    queryString: "*"
                },
                trimDuplicates: true // https://github.com/microsoftgraph/msgraph-sdk-java/issues/1315
            };
            const siteSearchRes = await getByGraphSearch(searchReqForSites);
            const locSearchItems = GraphSearchResponseMapper<ISiteSearchResponse>(siteSearchRes, _CONST.GraphSearch.SiteSearch.EntityType);
            const siteOptions = locSearchItems.map(item => {
                const parsedUrl = new URL(item.url);
                return {
                    key: item.url,
                    text: item.name + `(${parsedUrl.pathname})`
                }
            });
            setSiteFilterOptions(siteOptions);
        }
        getFilerValues().catch(error => console.log("init ~ error", error));
    }, []);

    const fileFolderOptions: IChoiceGroupOption[] = [
        { key: "BothFilesFolders" as FileFolderFilter, text: 'Both files & folders' },
        { key: 'OnlyFiles', text: 'Only files' },
        { key: 'Only Folders', text: 'Only folders' }
    ];

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
                        {/* <TextField
                            multiline
                            resizable={false}
                            value={filtreVal.siteUrl}
                            onChange={(e, newVal = '') => setFilterVal({ ...filtreVal, siteUrl: newVal })}
                            placeholder={`https://contoso.sharepoint.com/sites/...`}
                            description="Url of the site" /> */}
                        {siteFilterOptions?.length < 50 ? <>
                            <Dropdown
                                selectedKey={filtreVal.siteUrl}
                                options={siteFilterOptions}
                                onChange={(_e, op) => {
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
                            onChange={(_ex, checked) => {
                                setFilterVal({
                                    ...filtreVal,
                                    sharedType: checked ? [...filtreVal.sharedType, "Guest"] : filtreVal.sharedType.filter(val => val !== "Guest")
                                })
                            }}
                        />
                        <Checkbox label="Groups"
                            checked={filtreVal.sharedType.filter(val => val === "Group").length > 0}
                            onChange={(_ex, checked) => {
                                setFilterVal({
                                    ...filtreVal,
                                    sharedType: checked ? [...filtreVal.sharedType, "Group"] : filtreVal.sharedType.filter(val => val !== "Group")
                                })
                            }}
                        />
                        <Checkbox label="Member"
                            checked={filtreVal.sharedType.filter(val => val === "Member").length > 0}
                            onChange={(_ex, checked) => {
                                setFilterVal({
                                    ...filtreVal,
                                    sharedType: checked ? [...filtreVal.sharedType, "Member"] : filtreVal.sharedType.filter(val => val !== "Member")
                                })
                            }}
                        />
                    </div>

                </div>

                <div style={{ padding: "12px 0" }}>
                    <ChoiceGroup options={fileFolderOptions}
                        onChange={(_e, op) => {
                            setFilterVal({ ...filtreVal, fileFolder: op?.key as FileFolderFilter })
                        }} label="Item type"
                        required={true} />
                </div>
            </div>
        </Panel>

    </div>
}

export default FilterPanel;