import * as React from 'react';
import { Stack, SearchBox, PrimaryButton, FontIcon, Text } from '@fluentui/react';
import { useBoolean } from '@fluentui/react-hooks';
import { useContext } from 'react';
import { SharingWebPartContext } from '../../../hooks/SharingWebPartContext';
import { IPaginationFilterState } from '../SharingDetailedList';
import FilterPanel from './FilterPanel';


interface IFilterItemsProps {
    setPaginationFilterState: React.Dispatch<React.SetStateAction<IPaginationFilterState>>;
    paginationFilterState: IPaginationFilterState;
}

const FilterItems: React.FC<IFilterItemsProps> = (props): JSX.Element => {
    const governContext = useContext(SharingWebPartContext);
    const [isFilterPanelOpen, { setTrue: openFilterPanel, setFalse: dismissFilterPanel }] = useBoolean(false);

    return <>
        <div>
            <Stack horizontal horizontalAlign="space-between">
                <Stack.Item grow={3}>

                    <div style={{ maxWidth: "800px" }}>

                        <SearchBox
                            placeholder="Search..."
                            underlined={true}
                            onSearch={async (val: string) => {
                                props.setPaginationFilterState(prevState => ({
                                    ...prevState,
                                    searchKeyword: val,
                                    isRefreshData: true
                                }));
                            }}
                            onClear={async () => {
                                props.setPaginationFilterState(prevState => ({
                                    ...prevState,
                                    searchKeyword: "",
                                    isRefreshData: true
                                }));
                            }}
                        />
                    </div>

                </Stack.Item>
                <Stack horizontalAlign="end" style={{ marginLeft: 12 }}>
                    <PrimaryButton iconProps={{ iconName: 'Filter' }}
                        text="Filter" onClick={openFilterPanel} />

                    {isFilterPanelOpen &&
                        <FilterPanel
                            filterItem={props.paginationFilterState.filterVal}
                            isFilterPanelOpen={isFilterPanelOpen}
                            onDismissFilterPanel={async (newFilter) => {
                                dismissFilterPanel();
                                if (newFilter) {
                                    const updatedFilter: IPaginationFilterState = {
                                        ...props.paginationFilterState,
                                        filterVal: newFilter,
                                        currentPage: 1,
                                        isRefreshData: true
                                    };
                                    props.setPaginationFilterState(updatedFilter);
                                }
                            }}
                        />
                    }

                </Stack>
            </Stack>
        </div>

        <div style={{ paddingTop: "12px" }}>
            {props.paginationFilterState.searchMetadata &&
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    <Text variant="smallPlus">
                        <b>{props.paginationFilterState?.searchMetadata?.totalResults}</b> shared items found
                    </Text>
                    {governContext.webpartProperties.debugMode &&
                        <Text variant="smallPlus">
                            Query: <b>{props.paginationFilterState?.queryString}</b>
                        </Text>
                    }
                </div>
            }

        </div>

        <div style={{ padding: "12px 0" }}>
            {/* show filtered items here */}
            {!governContext.isTeams && <div>
                <FontIcon style={{ marginRight: 4, cursor: "pointer" }} aria-label="ClearFilter-Site" iconName="ClearFilter"
                    onClick={() => props.setPaginationFilterState(prevState => ({ ...prevState, filterVal: { ...prevState.filterVal, siteUrl: "" }, isRefreshData: true }))}
                />
                <Text variant="smallPlus">Site: <i>{props.paginationFilterState.filterVal.siteUrl || "All sites"}</i></Text>
            </div>}
            {props.paginationFilterState.filterVal.fileFolder !== "BothFilesFolders" && <div>
                <FontIcon style={{ marginRight: 4, cursor: "pointer" }} aria-label="ClearFilter-FileFolder" iconName="ClearFilter"
                    onClick={() => props.setPaginationFilterState(prevState => ({ ...prevState, filterVal: { ...prevState.filterVal, fileFolder: "BothFilesFolders" }, isRefreshData: true }))}
                />
                <Text variant="smallPlus">Type: <i>{props.paginationFilterState.filterVal.fileFolder}</i></Text>
            </div>}
        </div>
    </>;
};

export default FilterItems;