/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react'
import { Checkbox, ChoiceGroup, DefaultButton, IChoiceGroupOption, Label, Panel, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useContext, useState } from 'react';
import { SharedType } from '../../../../../common/model';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
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

    const { webpartContext } = useContext(SharingWebPartContext);
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: webpartContext.pageContext.web.absoluteUrl,
        msGraphClientFactory: webpartContext.msGraphClientFactory as any,
        spHttpClient: webpartContext.spHttpClient as any,
    };

    const [filtreVal, setFilterVal] = useState<IFilterItem>(props.filterItem);
    const [isSearchOneDrive, setIsSearchOneDrive] = useState<boolean>(filtreVal.siteUrl.indexOf("-my.sharepoint.com") > -1);

    // Define the options for the file/folder filter with the correct type for the key
    const fileFolderOptions: IChoiceGroupOption[] = [
        { key: "BothFilesFolders" as FileFolderFilter, text: 'Both files & folders' },
        { key: 'OnlyFiles' as FileFolderFilter, text: 'Only files' },
        { key: 'OnlyFolders' as FileFolderFilter, text: 'Only folders' }
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
                    <Toggle label="Search user ondrive"
                        checked={isSearchOneDrive}
                        onChange={(_e, checked) => {
                            setIsSearchOneDrive(checked ?? false);
                            if (checked) {
                                setFilterVal({ ...filtreVal, siteUrl: "" });
                            }
                        }} />

                    <div style={{ paddingLeft: "6px" }}>
                        <TextField
                            multiline
                            label='Site Url'
                            resizable={false}
                            disabled={isSearchOneDrive}
                            value={filtreVal.siteUrl}
                            onChange={(e, newVal = '') => setFilterVal({ ...filtreVal, siteUrl: newVal })}
                            placeholder={`https://contoso.sharepoint.com/sites/...`}
                            description="Url of the site" />

                        <PeoplePicker
                            context={peoplePickerContext}
                            titleText="OneDrive of the user"
                            personSelectionLimit={1}
                            disabled={!isSearchOneDrive}
                            searchTextLimit={2}
                            onChange={(items) => {
                                if (items.length > 0) {
                                    let selectedUserEmail = items[0].secondaryText ?? "";
                                    if (!selectedUserEmail) {
                                        selectedUserEmail = items[0].id ?? "";
                                        selectedUserEmail = selectedUserEmail.split("|")[2];
                                    }
                                    const locSelectedUser = selectedUserEmail.replace(/[^a-zA-Z0-9]/g, "_") ?? "";
                                    const tenantUrl = webpartContext.pageContext.web.absoluteUrl.split(".sharepoint.com")[0];
                                    const oneDriveUrl = `${tenantUrl}-my.sharepoint.com/personal/${locSelectedUser}`;
                                    setFilterVal({ ...filtreVal, siteUrl: oneDriveUrl });
                                } else {
                                    setFilterVal({ ...filtreVal, siteUrl: "" });
                                }
                            }
                            }
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={1000} />
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
                        selectedKey={filtreVal.fileFolder}
                        onChange={(_e, op) => {
                            setFilterVal({ ...filtreVal, fileFolder: op?.key as FileFolderFilter })
                        }} label="Item type" />
                </div>
            </div>
        </Panel>

    </div>
}

export default FilterPanel;