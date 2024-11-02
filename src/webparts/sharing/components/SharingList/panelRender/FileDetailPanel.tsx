

import * as React from 'react'
import { IFileSharingResponse } from '../../../../../common/model';
import { DefaultButton, Panel, PanelType, Pivot, PivotItem } from '@fluentui/react';
import FileActivityPanel from './FileActivityPanel';

interface IFileDetailPanelProps {
    file: IFileSharingResponse;
    isOpen: boolean;
    onDismiss(): void;
}


const FileDetailPanel: React.FC<IFileDetailPanelProps> = ({ file, isOpen, onDismiss }): JSX.Element => {

    return <div>
        <Panel
            isLightDismiss
            isOpen={isOpen}
            headerText={file.FileName}
            type={PanelType.medium}
            onDismiss={() => onDismiss()}
            onRenderFooterContent={() =>
                <div>
                    <DefaultButton onClick={() => onDismiss()}>Close</DefaultButton>
                </div>}
            isFooterAtBottom={true}
        >
            <div style={{ paddingTop: 32 }}>
                <Pivot aria-label="Basic Pivot Example">
                    <PivotItem headerText="Permission" itemIcon='ReminderPerson' >
                        <div>
                            <iframe
                                src={`${file.SiteUrl}/_layouts/15/sharedialog.aspx?listId=${file.ListId}&listItemId=${file.ListItemId}&clientId=sharePoint&mode=manageAccess&ma=0`} width="100%"
                                height={window.innerHeight - 260} />
                        </div>
                    </PivotItem>

                    <PivotItem
                        headerText="Activity"
                        itemIcon='LineChart'
                    >
                        <FileActivityPanel file={file} />
                    </PivotItem>
                </Pivot>
            </div>
        </Panel>
    </div>
}

export default FileDetailPanel;