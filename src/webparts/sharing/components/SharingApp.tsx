import * as React from 'react';
import { Text } from '@fluentui/react';
import SharingDetailedList from './SharingList/SharingDetailedList';
import { SharingWebPartContext } from '../hooks/SharingWebPartContext';


const SharingApp: React.FC = (): JSX.Element => {
    const governContext = React.useContext(SharingWebPartContext);
    const properties = governContext.webpartProperties;

    return <>
        {properties.webpartTitle &&
            <div>
                <Text variant={"xLarge"} style={{ paddingBottom: 24 }} nowrap block>
                    {properties.webpartTitle}
                </Text>

            </div>}

        <SharingDetailedList />
    </>;
};

export default SharingApp;