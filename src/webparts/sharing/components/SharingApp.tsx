import * as React from 'react';
import { Text } from '@fluentui/react';
import SharingDetailedList from './SharingList/SharingDetailedList';
import { SharingWebPartContext } from '../hooks/SharingWebPartContext';


const SharingApp: React.FC = (): JSX.Element => {
    const { webpartProperties } = React.useContext(SharingWebPartContext);

    return <>
        {webpartProperties.webpartTitle &&
            <div>
                <Text variant={"xLarge"} style={{ paddingBottom: 24 }} nowrap block>
                    {webpartProperties.webpartTitle}
                </Text>

            </div>}

        <SharingDetailedList />
    </>;
};

export default SharingApp;