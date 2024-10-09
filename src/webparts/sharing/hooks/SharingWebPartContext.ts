import { createContext } from 'react';
import { ISharingWebPartContext } from '../model';

export const SharingWebPartContext = createContext<
    ISharingWebPartContext
>(undefined);