

import * as React from 'react'
import { Icon } from '@fluentui/react';
import { FileIconType, getFileTypeIconProps } from '@fluentui/react-file-type-icons';

interface IFileExtentionColumnProps {
    ext: string;
}


const FileExtentionColumn: React.FC<IFileExtentionColumnProps> = ({ ext }): JSX.Element => {

    switch (ext) {
        case "folder": return <Icon {...getFileTypeIconProps({ type: FileIconType.documentsFolder, size: 16, imageFileType: 'png' })} />; break;
        default: return <Icon {...getFileTypeIconProps({ extension: `${ext}`, size: 16, imageFileType: 'png' })} />; break;
    }
}

export default FileExtentionColumn;