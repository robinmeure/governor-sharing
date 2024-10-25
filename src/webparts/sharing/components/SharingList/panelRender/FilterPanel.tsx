

import * as React from 'react'
import { Checkbox, DefaultButton, Label, Panel, PrimaryButton } from '@fluentui/react';
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

    const [filtreVal, setFilterVal] = useState<IFilterItem>(props.filterItem);
    console.log("FazLog ~ filtreVal:", filtreVal);

    return <div>
        <Panel
            headerText="Filter"
            isOpen={props.isFilterPanelOpen}
            onDismiss={() => props.onDismissFilterPanel()}
            closeButtonAriaLabel="Close"
            onRenderFooterContent={() =>
                <div>
                    <PrimaryButton onClick={() => {
                        setFilterVal(props.filterItem);
                        props.onDismissFilterPanel(filtreVal);
                    }}>Apply</PrimaryButton>
                    <DefaultButton style={{ marginLeft: 16 }} onClick={() => props.onDismissFilterPanel()}>Close</DefaultButton>
                </div>}
            isFooterAtBottom={true}
        >
            <div>

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