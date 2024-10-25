import * as React from 'react';
import SharingList from './SharingList/SharingList';


interface ISharingAppProps {

}

const SharingApp: React.FC<ISharingAppProps> = (): JSX.Element => {

    return <>
        <SharingList />
    </>;
};

export default SharingApp;