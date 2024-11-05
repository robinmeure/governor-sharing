

import * as React from 'react'
import { Text, Link } from '@fluentui/react';

interface ILinkColumnProps {
    label: string;
    url: string;
}


const LinkColumn: React.FC<ILinkColumnProps> = ({ label, url }): JSX.Element => {

    return <span><Text><Link target='_blank' data-interception="off" href={`${url}`}>{`${label}`}</Link></Text></span>
}

export default LinkColumn;