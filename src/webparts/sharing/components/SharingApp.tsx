import * as React from 'react';
import { useContext } from 'react';
import { SharingWebPartContext } from '../hooks/SharingWebPartContext';
import SharingList from './SharingList/SharingList';

const SharingApp: React.FC = (): JSX.Element => {

    const usdd = useContext(SharingWebPartContext);
    console.log("FazLog ~ usdd:", usdd.isTeams);

    return <>
        TEST
        <SharingList />
    </>;
};

export default SharingApp;